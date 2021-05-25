---
title: Outlook アドインで添付ファイルを取得する
description: アドインで添付ファイル API を使用して、添付ファイルに関する情報をリモート サービスに送信することができます。
ms.date: 01/14/2021
localization_priority: Normal
ms.openlocfilehash: db59ce44d2ed6f120503701479b705f13727130b
ms.sourcegitcommit: ecb24e32b32deb3e43daecd8d534e140460e0328
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/25/2021
ms.locfileid: "52639964"
---
# <a name="get-attachments-of-an-outlook-item-from-the-server"></a>サーバーから Outlook アイテムの添付ファイルを取得する

2 つの方法でOutlookアイテムの添付ファイルを取得できますが、使用するオプションはシナリオによって異なります。

1. リモート サービスに添付ファイル情報を送信します。

    アドインは添付ファイル API を使用して、添付ファイルに関する情報をリモート サービスに送信できます。 そうすれば、サービスは Exchange サーバーに直接アクセスして添付ファイルを取得できるようになります。

1. 要件セット 1.8 から利用できる [getAttachmentContentAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) API を使用します。 サポートされている形式: [AttachmentContentFormat](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat).

    この API は、EWS/REST が使用できない場合 (たとえば、Exchange サーバーの管理構成のため)、またはアドインが HTML または JavaScript で base64 コンテンツを直接使用する場合に便利です。 また、この API は、添付ファイルがまだ Exchange に同期されていない可能性がある作成シナリオで使用できます。詳細については `getAttachmentContentAsync` [、「Outlook](add-and-remove-attachments-to-an-item-in-a-compose-form.md)の作成フォームでアイテムの添付ファイルを管理する」を参照してください。

この記事では、最初のオプションについて詳しく説明します。 リモート サービスに添付ファイル情報を送信するには、次のプロパティと関数を使用します。

- [Office.context.mailbox.ewsUrl](/javascript/api/outlook/office.entities) プロパティ &ndash; メールボックスをホストしている Exchange サーバー上の Exchange Web サービス (EWS) の URL を指定します。サービスはこの URL を使用して、[ExchangeService.GetAttachments](/exchange/client-developer/exchange-web-services/how-to-get-attachments-by-using-ews-in-exchange) メソッドまたは [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) EWS 操作を呼び出します。

- [Office.context.mailbox.item.attachments](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) プロパティ &ndash; [AttachmentDetails](/javascript/api/outlook/office.attachmentdetails) オブジェクトの配列をアイテムの添付ファイルごとに 1 つ取得します。

- [Office.context.mailbox.getCallbackTokenAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) 関数 &ndash; メールボックスをホストする Exchange サーバーを非同期で呼び出し、添付ファイルの要求の認証のために Exchange サーバーに送り返すコールバック トークンを取得します。

## <a name="using-the-attachments-api"></a>添付ファイル API を使用する

添付ファイル API を使用してメールボックスから添付ファイルExchangeするには、次の手順を実行します。

1. 添付ファイルを含むメッセージまたは予定が表示されているときは、アドインを表示します。

1. Exchange サーバーからコールバック トークンを取得します。

1. コールバック トークンと添付ファイルの情報をリモート サービスに送信します。

1. `ExchangeService.GetAttachments` メソッドまたは `GetAttachment` 操作を使用して、Exchange サーバーから添付ファイルを取得します。

各手順については、以下のセクションで [Outlook-Add-in-JavaScript-GetAttachments](https://github.com/OfficeDev/Outlook-Add-in-JavaScript-GetAttachments) サンプルのコードを使用して詳しく説明します。

> [!NOTE]
> 以下の例に示すコードは、添付ファイルの情報を強調するために短縮されています。サンプルには、アドインをリモート サーバーで認証し、要求の状態を管理するためのコードも含まれています。

## <a name="get-a-callback-token"></a>コールバック トークンを取得する

[Office.context.mailbox](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md) オブジェクトは、Exchange サーバーで認証を行うためにリモート サーバーが使用できるトークンを取得するための `getCallbackTokenAsync` 関数を提供します。 次のコードは、コールバック トークンを取得するための非同期要求を起動するアドイン内の関数と、応答を取得するコールバック関数を示しています。 コールバック トークンは、次のセクションで定義されているサービス要求オブジェクトに保存されます。

```js
function getAttachmentToken() {
    if (serviceRequest.attachmentToken == "") {
        Office.context.mailbox.getCallbackTokenAsync(attachmentTokenCallback);
    }
}

function attachmentTokenCallback(asyncResult, userContext) {
    if (asyncResult.status === "succeeded") {
        // Cache the result from the server.
        serviceRequest.attachmentToken = asyncResult.value;
        serviceRequest.state = 3;
        testAttachments();
    } else {
        showToast("Error", "Could not get callback token: " + asyncResult.error.message);
    }
}
```

## <a name="send-attachment-information-to-the-remote-service"></a>添付ファイル情報をリモート サービスに送信する

アドインが呼び出すリモート サービスによって、サービスへの添付ファイル情報の送信方法に関する詳細が定義されます。この例では、リモート サービスは Visual Studio 2013 を使用して作成された Web API アプリケーションです。リモート サービスは、添付ファイルの情報が JSON オブジェクトに格納されていることを前提とします。次のコードは、添付ファイルの情報を格納するオブジェクトを初期化します。

```js
// Initialize a context object for the add-in.
//   Set the fields that are used on the request
//   object to default values.
 var serviceRequest = {
    attachmentToken: '',
    ewsUrl         : Office.context.mailbox.ewsUrl,
    attachments    : []
 };
```

<br/>

`Office.context.mailbox.item.attachments` プロパティには、アイテムの添付ファイルごとに存在する `AttachmentDetails` オブジェクトのコレクションが含まれています。 ほとんどの場合、アドインは `AttachmentDetails` オブジェクトの添付ファイル ID プロパティだけをリモート サービスに渡すことができます。 リモート サービスが添付ファイルについてより詳細な情報を必要とする場合、`AttachmentDetails` オブジェクトの全部あるいは一部を渡すことができます。 次のコードは、`AttachmentDetails` 配列全体を `serviceRequest` オブジェクトに配置し、リモート サービスに要求を送信するメソッドを定義しています。

```js
function makeServiceRequest() {
  // Format the attachment details for sending.
  for (var i = 0; i < mailbox.item.attachments.length; i++) {
    serviceRequest.attachments[i] = JSON.parse(JSON.stringify(mailbox.item.attachments[i]));
  }

  $.ajax({
    url: '../../api/Default',
    type: 'POST',
    data: JSON.stringify(serviceRequest),
    contentType: 'application/json;charset=utf-8'
  }).done(function (response) {
    if (!response.isError) {
      var names = "<h2>Attachments processed using " +
                    serviceRequest.service +
                    ": " +
                    response.attachmentsProcessed +
                    "</h2>";
      for (i = 0; i < response.attachmentNames.length; i++) {
        names += response.attachmentNames[i] + "<br />";
      }
      document.getElementById("names").innerHTML = names;
    } else {
      app.showNotification("Runtime error", response.message);
    }
  }).fail(function (status) {

  }).always(function () {
    $('.disable-while-sending').prop('disabled', false);
  })
}
```

## <a name="get-the-attachments-from-the-exchange-server"></a>Exchange サーバーから添付ファイルを取得する

リモート サービスは、 [GetAttachments](/exchange/client-developer/exchange-web-services/how-to-get-attachments-by-using-ews-in-exchange) EWS Managed API メソッドまたは [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) EWS 操作のいずれかを使用してサーバーから添付ファイルを取得できます。サービス アプリケーションは、JSON 文字列をサーバーで使用できる .NET Framework オブジェクトに逆シリアル化するために 2 つのオブジェクトを必要とします。次のコードに、逆シリアル化オブジェクトの定義を示します。

```cs
namespace AttachmentsSample
{
  public class AttachmentSampleServiceRequest
  {
    public string attachmentToken { get; set; }
    public string ewsUrl { get; set; }
    public string service { get; set; }
    public AttachmentDetails [] attachments { get; set; }
  }

  public class AttachmentDetails
  {
    public string attachmentType { get; set; }
    public string contentType { get; set; }
    public string id { get; set; }
    public bool isInline { get; set; }
    public string name { get; set; }
    public int size { get; set; }
  }
}
```

### <a name="use-the-ews-managed-api-to-get-the-attachments"></a>EWS Managed API を使用して添付ファイルを取得する

リモート サービスで [EWS Managed API](/exchange/client-developer/web-service-reference/ews-managed-api-reference-for-exchange) を使用する場合は、 [GetAttachments](/exchange/client-developer/exchange-web-services/how-to-get-attachments-by-using-ews-in-exchange) メソッドを使用できます。このメソッドは、添付ファイルを取得するための EWS SOAP 要求を作成、送信、および受信します。EWS Managed API を使用すると、必要なコード行が少なく、EWS の呼び出しを行うための直感的なインターフェイスが提供されるため、この API の使用をお勧めします。次のコードは、1 回の要求ですべての添付ファイルを取得し、処理された添付ファイルの数と名前を返します。

```cs
private AttachmentSampleServiceResponse GetAtttachmentsFromExchangeServerUsingEWSManagedApi(AttachmentSampleServiceRequest request)
{
  var attachmentsProcessedCount = 0;
  var attachmentNames = new List<string>();

  // Create an ExchangeService object, set the credentials and the EWS URL.
  ExchangeService service = new ExchangeService();
  service.Credentials = new OAuthCredentials(request.attachmentToken);
  service.Url = new Uri(request.ewsUrl);

  var attachmentIds = new List<string>();

  foreach (AttachmentDetails attachment in request.attachments)
  {
    attachmentIds.Add(attachment.id);
  }

  // Call the GetAttachments method to retrieve the attachments on the message.
  // This method results in a GetAttachments EWS SOAP request and response
  // from the Exchange server.
  var getAttachmentsResponse =
    service.GetAttachments(attachmentIds.ToArray(),
                            null,
                            new PropertySet(BasePropertySet.FirstClassProperties,
                                            ItemSchema.MimeContent));

  if (getAttachmentsResponse.OverallResult == ServiceResult.Success)
  {
    foreach (var attachmentResponse in getAttachmentsResponse)
    {
      attachmentNames.Add(attachmentResponse.Attachment.Name);

      // Write the content of each attachment to a stream.
      if (attachmentResponse.Attachment is FileAttachment)
      {
        FileAttachment fileAttachment = attachmentResponse.Attachment as FileAttachment;
        Stream s = new MemoryStream(fileAttachment.Content);
        // Process the contents of the attachment here.
      }

      if (attachmentResponse.Attachment is ItemAttachment)
      {
        ItemAttachment itemAttachment = attachmentResponse.Attachment as ItemAttachment;
        Stream s = new MemoryStream(itemAttachment.Item.MimeContent.Content);
        // Process the contents of the attachment here.
      }

      attachmentsProcessedCount++;
    }
  }

  // Return the names and number of attachments processed for display
  // in the add-in UI.
  var response = new AttachmentSampleServiceResponse();
  response.attachmentNames = attachmentNames.ToArray();
  response.attachmentsProcessed = attachmentsProcessedCount;

  return response;
}
```

### <a name="use-ews-to-get-the-attachments"></a>EWS を使用して添付ファイルを取得する

リモート サービスで EWS を使用している場合、Exchange サーバーから添付ファイルを取得するために [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) SOAP 要求を作成する必要があります。 次のコードは、SOAP 要求を提供する文字列を返します。 リモート サービスは、添付ファイル用の添付ファイル ID を文字列に挿入するために `String.Format` メソッドを使用します。


```cs
private const string GetAttachmentSoapRequest =
@"<?xml version=""1.0"" encoding=""utf-8""?>
<soap:Envelope xmlns:xsi=""https://www.w3.org/2001/XMLSchema-instance""
xmlns:xsd=""https://www.w3.org/2001/XMLSchema""
xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/""
xmlns:t=""http://schemas.microsoft.com/exchange/services/2006/types"">
<soap:Header>
<t:RequestServerVersion Version=""Exchange2013"" />
</soap:Header>
  <soap:Body>
    <GetAttachment xmlns=""http://schemas.microsoft.com/exchange/services/2006/messages""
    xmlns:t=""http://schemas.microsoft.com/exchange/services/2006/types"">
      <AttachmentShape/>
      <AttachmentIds>
        <t:AttachmentId Id=""{0}""/>
      </AttachmentIds>
    </GetAttachment>
  </soap:Body>
</soap:Envelope>";
```

<br/>

最後に、次のメソッドでは、Exchange サーバーから添付ファイルを取得するために EWS の `GetAttachment` 要求を使用するという作業を行っています。 この実装では、各添付ファイルに対して個別の要求を出し、処理された添付ファイルの数を返します。 各応答は、次に定義されているような、それぞれ別々の `ProcessXmlResponse` メソッドで処理されます。

```cs
private AttachmentSampleServiceResponse GetAttachmentsFromExchangeServerUsingEWS(AttachmentSampleServiceRequest request)
{
  var attachmentsProcessedCount = 0;
  var attachmentNames = new List<string>();

  foreach (var attachment in request.attachments)
  {
    // Prepare a web request object.
    HttpWebRequest webRequest = WebRequest.CreateHttp(request.ewsUrl);
    webRequest.Headers.Add("Authorization",
      string.Format("Bearer {0}", request.attachmentToken));
    webRequest.PreAuthenticate = true;
    webRequest.AllowAutoRedirect = false;
    webRequest.Method = "POST";
    webRequest.ContentType = "text/xml; charset=utf-8";

    // Construct the SOAP message for the GetAttachment operation.
    byte[] bodyBytes = Encoding.UTF8.GetBytes(
      string.Format(GetAttachmentSoapRequest, attachment.id));
    webRequest.ContentLength = bodyBytes.Length;

    Stream requestStream = webRequest.GetRequestStream();
    requestStream.Write(bodyBytes, 0, bodyBytes.Length);
    requestStream.Close();

    // Make the request to the Exchange server and get the response.
    HttpWebResponse webResponse = (HttpWebResponse)webRequest.GetResponse();

    // If the response is okay, create an XML document from the response
    // and process the request.
    if (webResponse.StatusCode == HttpStatusCode.OK)
    {
      var responseStream = webResponse.GetResponseStream();

      var responseEnvelope = XElement.Load(responseStream);

      // After creating a memory stream containing the contents of the
      // attachment, this method writes the XML document to the trace output.
      // Your service would perform it's processing here.
      if (responseEnvelope != null)
      {
        var processResult = ProcessXmlResponse(responseEnvelope);
        attachmentNames.Add(string.Format("{0} {1}", attachment.name, processResult));

      }

      // Close the response stream.
      responseStream.Close();
      webResponse.Close();

    }
    // If the response is not OK, return an error message for the
    // attachment.
    else
    {
      var errorString = string.Format("Attachment \"{0}\" could not be processed. " +
        "Error message: {1}.", attachment.name, webResponse.StatusDescription);
      attachmentNames.Add(errorString);
    }
    attachmentsProcessedCount++;
  }

  // Return the names and number of attachments processed for display
  // in the add-in UI.
  var response = new AttachmentSampleServiceResponse();
  response.attachmentNames = attachmentNames.ToArray();
  response.attachmentsProcessed = attachmentsProcessedCount;

  return response;
}
```

<br/>

各 `GetAttachment` 操作からのそれぞれの応答は、`ProcessXmlResponse` メソッドに送信されます。 このメソッドは応答にエラーがないか確認します。 エラーが見つからなかった場合、添付ファイルや添付アイテムを処理します。 `ProcessXmlResponse` メソッドが添付ファイル処理作業の大部分を行います。

```cs
// This method processes the response from the Exchange server.
// In your application the bulk of the processing occurs here.
private string ProcessXmlResponse(XElement responseEnvelope)
{
  // First, check the response for web service errors.
  var errorCodes = from errorCode in responseEnvelope.Descendants
                    ("{http://schemas.microsoft.com/exchange/services/2006/messages}ResponseCode")
                    select errorCode;
  // Return the first error code found.
  foreach (var errorCode in errorCodes)
  {
    if (errorCode.Value != "NoError")
    {
      return string.Format("Could not process result. Error: {0}", errorCode.Value);
    }
  }

  // No errors found, proceed with processing the content.
  // First, get and process file attachments.
  var fileAttachments = from fileAttachment in responseEnvelope.Descendants
                    ("{http://schemas.microsoft.com/exchange/services/2006/types}FileAttachment")
                        select fileAttachment;
  foreach(var fileAttachment in fileAttachments)
  {
    var fileContent = fileAttachment.Element("{http://schemas.microsoft.com/exchange/services/2006/types}Content");
    var fileData = System.Convert.FromBase64String(fileContent.Value);
    var s = new MemoryStream(fileData);
    // Process the file attachment here.
  }

  // Second, get and process item attachments.
  var itemAttachments = from itemAttachment in responseEnvelope.Descendants
                        ("{http://schemas.microsoft.com/exchange/services/2006/types}ItemAttachment")
                        select itemAttachment;
  foreach(var itemAttachment in itemAttachments)
  {
    var message = itemAttachment.Element("{http://schemas.microsoft.com/exchange/services/2006/types}Message");
    if (message != null)
    {
      // Process a message here.
      break;
    }
    var calendarItem = itemAttachment.Element("{http://schemas.microsoft.com/exchange/services/2006/types}CalendarItem");
    if (calendarItem != null)
    {
      // Process calendar item here.
      break;
    }
    var contact = itemAttachment.Element("{http://schemas.microsoft.com/exchange/services/2006/types}Contact");
    if (contact != null)
    {
      // Process contact here.
      break;
    }
    var task = itemAttachment.Element("{http://schemas.microsoft.com/exchange/services/2006/types}Tontact");
    if (task != null)
    {
      // Process task here.
      break;
    }
    var meetingMessage = itemAttachment.Element("{http://schemas.microsoft.com/exchange/services/2006/types}MeetingMessage");
    if (meetingMessage != null)
    {
      // Process meeting message here.
      break;
    }
    var meetingRequest = itemAttachment.Element("{http://schemas.microsoft.com/exchange/services/2006/types}MeetingRequest");
    if (meetingRequest != null)
    {
      // Process meeting request here.
      break;
    }
    var meetingResponse = itemAttachment.Element("{http://schemas.microsoft.com/exchange/services/2006/types}MeetingResponse");
    if (meetingResponse != null)
    {
      // Process meeting response here.
      break;
    }
    var meetingCancellation = itemAttachment.Element("{http://schemas.microsoft.com/exchange/services/2006/types}MeetingCancellation");
    if (meetingCancellation != null)
    {
      // Process meeting cancellation here.
      break;
    }
  }

  return string.Empty;
}
```

## <a name="see-also"></a>関連項目

- [閲覧フォーム用の Outlook アドインを作成する](read-scenario.md)
- 
  [Exchange の EWS Managed API、EWS、および Web サービスについて学ぶ](/exchange/client-developer/exchange-web-services/explore-the-ews-managed-api-ews-and-web-services-in-exchange)
- [EWS マネージ API クライアント アプリケーションの概要](/exchange/client-developer/exchange-web-services/get-started-with-ews-managed-api-client-applications)
- [Outlookアドイン SSO](https://github.com/OfficeDev/Outlook-Add-in-SSO)
