using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Bot.Builder;
using Microsoft.Bot.Builder.Teams;
using Microsoft.Bot.Schema;
using Microsoft.Extensions.Configuration;
using Microsoft.Graph;
using Microsoft.Identity.Client;

namespace EchoBot.Bots
{
    /// <summary>
    /// 这个类用来进行身份验证
    /// </summary>
    public class SimpleAuth : IAuthenticationProvider
    {
        // 读取配置文件（或云平台的配置）
        private IConfiguration _configuration;
        public SimpleAuth(IConfiguration configuration)
        {
            this._configuration = configuration;
        }
        public async Task AuthenticateRequestAsync(HttpRequestMessage request)
        {
            // 这个应用是指后台应用，不需要用户参与授权。请注意，使用密码的方式并不是很推荐的，在生产环境，尽量用证书。
            IConfidentialClientApplication app = ConfidentialClientApplicationBuilder
                .Create(_configuration.GetValue<string>("MicrosoftAppId"))
                .WithTenantId(_configuration.GetValue<string>("TenantId"))
                .WithClientSecret(_configuration.GetValue<string>("MicrosoftAppPassword"))
                .Build();

            // 获取访问凭据。请注意，生产环境下，这里需要考虑缓存。
            var token = await app.AcquireTokenForClient(new[] { "https://graph.microsoft.com/.default" }).ExecuteAsync();
            // 把AccessToken 附加的请求的头部中去
            request.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", token.AccessToken);
        }
    }
    public class EchoBot : TeamsActivityHandler
    {
        private IConfiguration _configuration;
        public EchoBot(IConfiguration config)
        {
            this._configuration = config;
        }

        protected override async Task OnMessageActivityAsync(ITurnContext<IMessageActivity> turnContext, CancellationToken cancellationToken)
        {
            // 读取当前群聊中的用户列表
            var members = await TeamsInfo.GetMembersAsync(turnContext);
            // 准备用来访问Microsoft Graph的本地代理
            GraphServiceClient graphClient = new GraphServiceClient(new SimpleAuth(_configuration));
            // 为每个用户发送一个通知，这里解析得到他们的AadObjectId，用来发送
            members.Select(_ => _.AadObjectId).AsParallel().ForAll(async (_) =>
            {
                // 以下代码，其实你可以在官网找到，并且简单地做些修改即可 
                // https://docs.microsoft.com/zh-cn/graph/api/chat-sendactivitynotification?view=graph-rest-1.0&tabs=http#example-1-notify-a-user-about-a-task-created-in-a-chat
                var topic = new TeamworkActivityTopic
                {
                    Source = TeamworkActivityTopicSource.EntityUrl,
                    Value = $"https://graph.microsoft.com/beta/me/chats/{turnContext.Activity.Conversation.Id}/messages/{turnContext.Activity.Id}"
                };
                // 这个是通知的自定义模板（在manifest.json文件中要定义）
                var activityType = "metionall";
                // 预览文字
                var previewText = new ItemBody
                {
                    Content = "有人在群聊中提到你了，请点击查看"
                };
                // 收件人的id
                var recipient = new AadUserNotificationRecipient
                {
                    UserId = _
                };
                // 替换掉模板中的值
                var templateParameters = new List<Microsoft.Graph.KeyValuePair>()
                {
                    new Microsoft.Graph.KeyValuePair
                    {
                        Name = "from",
                        Value = turnContext.Activity.From.Name
                    },
                    new Microsoft.Graph.KeyValuePair
                    {
                        Name ="message",
                        Value =turnContext.Activity.RemoveMentionText(turnContext.Activity.Recipient.Id)
                    }
                };
                // 调用接口发送通知
                await graphClient.Chats[turnContext.Activity.Conversation.Id]
                    .SendActivityNotification(topic, activityType, null, previewText, templateParameters, recipient)
                    .Request()
                    .PostAsync();
            });

        }
    }
}
