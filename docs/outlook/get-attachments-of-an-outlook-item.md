---
title: Outlook アドインで添付ファイルを取得する
description: アドインで添付ファイル API を使用して、添付ファイルに関する情報をリモート サービスに送信することができます。
ms.date: 08/20/2020
localization_priority: Normal
ms.openlocfilehash: 57191820e27bc78431d0a7c97ffd6b8f23e75f4b
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/28/2020
ms.locfileid: "47293913"
---
# <a name="get-attachments-of-an-outlook-item-from-the-server"></a><span data-ttu-id="515eb-103">サーバーから Outlook アイテムの添付ファイルを取得する</span><span class="sxs-lookup"><span data-stu-id="515eb-103">Get attachments of an Outlook item from the server</span></span>

<span data-ttu-id="515eb-p101">Outlook アドインは、サーバーで実行されているリモート サービスに選択したアイテムの添付ファイルを直接渡すことができません。その代わりに、添付ファイル API を使用して、添付ファイルに関する情報をリモート サービスに送信できます。そうすれば、サービスは Exchange サーバーに直接アクセスして添付ファイルを取得できるようになります。</span><span class="sxs-lookup"><span data-stu-id="515eb-p101">An Outlook add-in cannot pass the attachments of a selected item directly to the remote service that runs on your server. Instead, the add-in can use the attachments API to send information about the attachments to the remote service. The service can then contact the Exchange server directly to retrieve the attachments.</span></span>

<span data-ttu-id="515eb-107">添付ファイルの情報をリモート サービスに送信するには、次のプロパティと関数を使用します。</span><span class="sxs-lookup"><span data-stu-id="515eb-107">To send attachment information to the remote service, you use the following properties and function:</span></span>

- <span data-ttu-id="515eb-p102">[Office.context.mailbox.ewsUrl](/javascript/api/outlook/office.entities) プロパティ &ndash; メールボックスをホストしている Exchange サーバー上の Exchange Web サービス (EWS) の URL を指定します。サービスはこの URL を使用して、[ExchangeService.GetAttachments](/exchange/client-developer/exchange-web-services/how-to-get-attachments-by-using-ews-in-exchange) メソッドまたは [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) EWS 操作を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="515eb-p102">[Office.context.mailbox.ewsUrl](/javascript/api/outlook/office.entities) property &ndash; Provides the URL of Exchange Web Services (EWS) on the Exchange server that hosts the mailbox. Your service uses this URL to call the [ExchangeService.GetAttachments](/exchange/client-developer/exchange-web-services/how-to-get-attachments-by-using-ews-in-exchange) method, or the [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) EWS operation.</span></span>

- <span data-ttu-id="515eb-110">[Office.context.mailbox.item.attachments](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) プロパティ &ndash; [AttachmentDetails](/javascript/api/outlook/office.attachmentdetails) オブジェクトの配列をアイテムの添付ファイルごとに 1 つ取得します。</span><span class="sxs-lookup"><span data-stu-id="515eb-110">[Office.context.mailbox.item.attachments](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) property &ndash; Gets an array of [AttachmentDetails](/javascript/api/outlook/office.attachmentdetails) objects, one for each attachment to the item.</span></span>

- <span data-ttu-id="515eb-111">[Office.context.mailbox.getCallbackTokenAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) 関数 &ndash; メールボックスをホストする Exchange サーバーを非同期で呼び出し、添付ファイルの要求の認証のために Exchange サーバーに送り返すコールバック トークンを取得します。</span><span class="sxs-lookup"><span data-stu-id="515eb-111">[Office.context.mailbox.getCallbackTokenAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) function &ndash; Makes an asynchronous call to the Exchange server that hosts the mailbox to get a callback token that the server sends back to the Exchange server to authenticate a request for an attachment.</span></span>

## <a name="using-the-attachments-api"></a><span data-ttu-id="515eb-112">添付ファイル API を使用する</span><span class="sxs-lookup"><span data-stu-id="515eb-112">Using the attachments API</span></span>

<span data-ttu-id="515eb-113">添付ファイル API を使用して Exchange メールボックスから添付ファイルを取得するには、次の手順を実行します。</span><span class="sxs-lookup"><span data-stu-id="515eb-113">To use the attachments API to get attachments from an Exchange mailbox, perform the following steps:</span></span>

1. <span data-ttu-id="515eb-114">添付ファイルを含むメッセージまたは予定が表示されているときは、アドインを表示します。</span><span class="sxs-lookup"><span data-stu-id="515eb-114">Show the add-in when the user is viewing a message or appointment that contains an attachment.</span></span>

1. <span data-ttu-id="515eb-115">Exchange サーバーからコールバック トークンを取得します。</span><span class="sxs-lookup"><span data-stu-id="515eb-115">Get the callback token from the Exchange server.</span></span>

1. <span data-ttu-id="515eb-116">コールバック トークンと添付ファイルの情報をリモート サービスに送信します。</span><span class="sxs-lookup"><span data-stu-id="515eb-116">Send the callback token and attachment information to the remote service.</span></span>

1. <span data-ttu-id="515eb-117">`ExchangeService.GetAttachments` メソッドまたは `GetAttachment` 操作を使用して、Exchange サーバーから添付ファイルを取得します。</span><span class="sxs-lookup"><span data-stu-id="515eb-117">Get the attachments from the Exchange server by using the `ExchangeService.GetAttachments` method or the `GetAttachment` operation.</span></span>

<span data-ttu-id="515eb-118">各手順については、以下のセクションで [Outlook-Add-in-JavaScript-GetAttachments](https://github.com/OfficeDev/Outlook-Add-in-JavaScript-GetAttachments) サンプルのコードを使用して詳しく説明します。</span><span class="sxs-lookup"><span data-stu-id="515eb-118">Each of these steps is covered in detail in the following sections using code from the [Outlook-Add-in-JavaScript-GetAttachments](https://github.com/OfficeDev/Outlook-Add-in-JavaScript-GetAttachments) sample.</span></span>

> [!NOTE]
> <span data-ttu-id="515eb-p103">以下の例に示すコードは、添付ファイルの情報を強調するために短縮されています。サンプルには、アドインをリモート サーバーで認証し、要求の状態を管理するためのコードも含まれています。</span><span class="sxs-lookup"><span data-stu-id="515eb-p103">The code in these examples has been shortened to emphasize the attachment information. The sample contains additional code for authenticating the add-in with the remote server and managing the state of the request.</span></span>

## <a name="get-a-callback-token"></a><span data-ttu-id="515eb-121">コールバック トークンを取得する</span><span class="sxs-lookup"><span data-stu-id="515eb-121">Get a callback token</span></span>

<span data-ttu-id="515eb-122">[Office.context.mailbox](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md) オブジェクトは、Exchange サーバーで認証を行うためにリモート サーバーが使用できるトークンを取得するための `getCallbackTokenAsync` 関数を提供します。</span><span class="sxs-lookup"><span data-stu-id="515eb-122">The [Office.context.mailbox](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md) object provides the `getCallbackTokenAsync` function to get a token that the remote server can use to authenticate with the Exchange server.</span></span> <span data-ttu-id="515eb-123">次のコードは、コールバック トークンを取得するための非同期要求を起動するアドイン内の関数と、応答を取得するコールバック関数を示しています。</span><span class="sxs-lookup"><span data-stu-id="515eb-123">The following code shows a function in an add-in that starts the asynchronous request to get the callback token, and the callback function that gets the response.</span></span> <span data-ttu-id="515eb-124">コールバック トークンは、次のセクションで定義されているサービス要求オブジェクトに保存されます。</span><span class="sxs-lookup"><span data-stu-id="515eb-124">The callback token is stored in the service request object that is defined in the next section.</span></span>

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

## <a name="send-attachment-information-to-the-remote-service"></a><span data-ttu-id="515eb-125">添付ファイル情報をリモート サービスに送信する</span><span class="sxs-lookup"><span data-stu-id="515eb-125">Send attachment information to the remote service</span></span>

<span data-ttu-id="515eb-p105">アドインが呼び出すリモート サービスによって、サービスへの添付ファイル情報の送信方法に関する詳細が定義されます。この例では、リモート サービスは Visual Studio 2013 を使用して作成された Web API アプリケーションです。リモート サービスは、添付ファイルの情報が JSON オブジェクトに格納されていることを前提とします。次のコードは、添付ファイルの情報を格納するオブジェクトを初期化します。</span><span class="sxs-lookup"><span data-stu-id="515eb-p105">The remote service that your add-in calls defines the specifics of how you should send the attachment information to the service. In this example, the remote service is a Web API application created by using Visual Studio 2013. The remote service expects the attachment information in a JSON object. The following code initializes an object that contains the attachment information.</span></span>

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

<span data-ttu-id="515eb-130">`Office.context.mailbox.item.attachments` プロパティには、アイテムの添付ファイルごとに存在する `AttachmentDetails` オブジェクトのコレクションが含まれています。</span><span class="sxs-lookup"><span data-stu-id="515eb-130">The `Office.context.mailbox.item.attachments` property contains a collection of `AttachmentDetails` objects, one for each attachment to the item.</span></span> <span data-ttu-id="515eb-131">ほとんどの場合、アドインは `AttachmentDetails` オブジェクトの添付ファイル ID プロパティだけをリモート サービスに渡すことができます。</span><span class="sxs-lookup"><span data-stu-id="515eb-131">In most cases, the add-in can pass just the attachment ID property of an `AttachmentDetails` object to the remote service.</span></span> <span data-ttu-id="515eb-132">リモート サービスが添付ファイルについてより詳細な情報を必要とする場合、`AttachmentDetails` オブジェクトの全部あるいは一部を渡すことができます。</span><span class="sxs-lookup"><span data-stu-id="515eb-132">If the remote service needs more details about the attachment, you can pass all or part of the `AttachmentDetails` object.</span></span> <span data-ttu-id="515eb-133">次のコードは、`AttachmentDetails` 配列全体を `serviceRequest` オブジェクトに配置し、リモート サービスに要求を送信するメソッドを定義しています。</span><span class="sxs-lookup"><span data-stu-id="515eb-133">The following code defines a method that puts the entire `AttachmentDetails` array in the `serviceRequest` object and sends a request to the remote service.</span></span>

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

## <a name="get-the-attachments-from-the-exchange-server"></a><span data-ttu-id="515eb-134">Exchange サーバーから添付ファイルを取得する</span><span class="sxs-lookup"><span data-stu-id="515eb-134">Get the attachments from the Exchange server</span></span>

<span data-ttu-id="515eb-p107">リモート サービスは、 [GetAttachments](/exchange/client-developer/exchange-web-services/how-to-get-attachments-by-using-ews-in-exchange) EWS Managed API メソッドまたは [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) EWS 操作のいずれかを使用してサーバーから添付ファイルを取得できます。サービス アプリケーションは、JSON 文字列をサーバーで使用できる .NET Framework オブジェクトに逆シリアル化するために 2 つのオブジェクトを必要とします。次のコードに、逆シリアル化オブジェクトの定義を示します。</span><span class="sxs-lookup"><span data-stu-id="515eb-p107">Your remote service can use either the [GetAttachments](/exchange/client-developer/exchange-web-services/how-to-get-attachments-by-using-ews-in-exchange) EWS Managed API method or the [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) EWS operation to retrieve attachments from the server. The service application needs two objects to deserialize the JSON string into .NET Framework objects that can be used on the server. The following code shows the definitions of the deserialization objects.</span></span>

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

### <a name="use-the-ews-managed-api-to-get-the-attachments"></a><span data-ttu-id="515eb-138">EWS Managed API を使用して添付ファイルを取得する</span><span class="sxs-lookup"><span data-stu-id="515eb-138">Use the EWS Managed API to get the attachments</span></span>

<span data-ttu-id="515eb-p108">リモート サービスで [EWS Managed API](https://go.microsoft.com/fwlink/?LinkID=255472) を使用する場合は、 [GetAttachments](/exchange/client-developer/exchange-web-services/how-to-get-attachments-by-using-ews-in-exchange) メソッドを使用できます。このメソッドは、添付ファイルを取得するための EWS SOAP 要求を作成、送信、および受信します。EWS Managed API を使用すると、必要なコード行が少なく、EWS の呼び出しを行うための直感的なインターフェイスが提供されるため、この API の使用をお勧めします。次のコードは、1 回の要求ですべての添付ファイルを取得し、処理された添付ファイルの数と名前を返します。</span><span class="sxs-lookup"><span data-stu-id="515eb-p108">If you use the [EWS Managed API](https://go.microsoft.com/fwlink/?LinkID=255472) in your remote service, you can use the [GetAttachments](/exchange/client-developer/exchange-web-services/how-to-get-attachments-by-using-ews-in-exchange) method, which will construct, send, and receive an EWS SOAP request to get the attachments. We recommend that you use the EWS Managed API because it requires fewer lines of code and provides a more intuitive interface for making calls to EWS. The following code makes one request to retrieve all the attachments, and returns the count and names of the attachments processed.</span></span>

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

### <a name="use-ews-to-get-the-attachments"></a><span data-ttu-id="515eb-142">EWS を使用して添付ファイルを取得する</span><span class="sxs-lookup"><span data-stu-id="515eb-142">Use EWS to get the attachments</span></span>

<span data-ttu-id="515eb-143">リモート サービスで EWS を使用している場合、Exchange サーバーから添付ファイルを取得するために [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) SOAP 要求を作成する必要があります。</span><span class="sxs-lookup"><span data-stu-id="515eb-143">If you use EWS in your remote service, you need to construct a [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) SOAP request to get the attachments from the Exchange server.</span></span> <span data-ttu-id="515eb-144">次のコードは、SOAP 要求を提供する文字列を返します。</span><span class="sxs-lookup"><span data-stu-id="515eb-144">The following code returns a string that provides the SOAP request.</span></span> <span data-ttu-id="515eb-145">リモート サービスは、添付ファイル用の添付ファイル ID を文字列に挿入するために `String.Format` メソッドを使用します。</span><span class="sxs-lookup"><span data-stu-id="515eb-145">The remote service uses the `String.Format` method to insert the attachment ID for an attachment into the string.</span></span>


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

<span data-ttu-id="515eb-146">最後に、次のメソッドでは、Exchange サーバーから添付ファイルを取得するために EWS の `GetAttachment` 要求を使用するという作業を行っています。</span><span class="sxs-lookup"><span data-stu-id="515eb-146">Finally, the following method does the work of using an EWS `GetAttachment` request to get the attachments from the Exchange server.</span></span> <span data-ttu-id="515eb-147">この実装では、各添付ファイルに対して個別の要求を出し、処理された添付ファイルの数を返します。</span><span class="sxs-lookup"><span data-stu-id="515eb-147">This implementation makes an individual request for each attachment, and returns the count of attachments processed.</span></span> <span data-ttu-id="515eb-148">各応答は、次に定義されているような、それぞれ別々の `ProcessXmlResponse` メソッドで処理されます。</span><span class="sxs-lookup"><span data-stu-id="515eb-148">Each response is processed in a separate `ProcessXmlResponse` method, defined next.</span></span>

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

<span data-ttu-id="515eb-149">各 `GetAttachment` 操作からのそれぞれの応答は、`ProcessXmlResponse` メソッドに送信されます。</span><span class="sxs-lookup"><span data-stu-id="515eb-149">Each response from the `GetAttachment` operation is sent to the `ProcessXmlResponse` method.</span></span> <span data-ttu-id="515eb-150">このメソッドは応答にエラーがないか確認します。</span><span class="sxs-lookup"><span data-stu-id="515eb-150">This method checks the response for errors.</span></span> <span data-ttu-id="515eb-151">エラーが見つからなかった場合、添付ファイルや添付アイテムを処理します。</span><span class="sxs-lookup"><span data-stu-id="515eb-151">If it doesn't find any errors, it processes file attachments and item attachments.</span></span> <span data-ttu-id="515eb-152">`ProcessXmlResponse` メソッドが添付ファイル処理作業の大部分を行います。</span><span class="sxs-lookup"><span data-stu-id="515eb-152">The `ProcessXmlResponse` method performs the bulk of the work to process the attachment.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="515eb-153">関連項目</span><span class="sxs-lookup"><span data-stu-id="515eb-153">See also</span></span>

- [<span data-ttu-id="515eb-154">閲覧フォーム用の Outlook アドインを作成する</span><span class="sxs-lookup"><span data-stu-id="515eb-154">Create Outlook add-ins for read forms</span></span>](read-scenario.md)
- <span data-ttu-id="515eb-155">
  [Exchange の EWS Managed API、EWS、および Web サービスについて学ぶ](/exchange/client-developer/exchange-web-services/explore-the-ews-managed-api-ews-and-web-services-in-exchange)</span><span class="sxs-lookup"><span data-stu-id="515eb-155">[Explore the EWS Managed API, EWS, and web services in Exchange](/exchange/client-developer/exchange-web-services/explore-the-ews-managed-api-ews-and-web-services-in-exchange)</span></span>
- [<span data-ttu-id="515eb-156">EWS マネージ API クライアント アプリケーションの概要</span><span class="sxs-lookup"><span data-stu-id="515eb-156">Get started with EWS Managed API client applications</span></span>](/exchange/client-developer/exchange-web-services/get-started-with-ews-managed-api-client-applications)
- [<span data-ttu-id="515eb-157">Outlook アドイン SSO</span><span class="sxs-lookup"><span data-stu-id="515eb-157">Outlook Add-in SSO</span></span>](https://github.com/OfficeDev/Outlook-Add-in-SSO)
