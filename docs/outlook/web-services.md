---
title: Outlook アドインから Exchange Web サービス (EWS) を使用する
description: Outlook アドインが Exchange Web サービスに情報を要求する方法の例を示します。
ms.date: 04/28/2020
localization_priority: Normal
ms.openlocfilehash: a8299b3e96db48c296fe0e61b36668a788fb8799
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/28/2020
ms.locfileid: "47292483"
---
# <a name="call-web-services-from-an-outlook-add-in"></a><span data-ttu-id="97675-103">Outlook アドインから Web サービスを呼び出す</span><span class="sxs-lookup"><span data-stu-id="97675-103">Call web services from an Outlook add-in</span></span>

<span data-ttu-id="97675-p101">アドインは、Exchange Server 2013 が実行されているコンピューター、アドインの UI のソースの場所を提供するサーバー上で利用できる Web サービス、またはインターネット上で利用できる Web サービスから Exchange Web サービス (EWS) を使用できます。この記事では、Outlook アドインからどのように EWS の情報を要求できるかを示す例を説明します。</span><span class="sxs-lookup"><span data-stu-id="97675-p101">Your add-in can use Exchange Web Services (EWS) from a computer that is running Exchange Server 2013, a web service that is available on the server that provides the source location for the add-in's UI, or a web service that is available on the Internet. This article provides an example that shows how an Outlook add-in can request information from EWS.</span></span>

<span data-ttu-id="97675-p102">Web サービスを呼び出す方法は、Web サービスは配置された場所によって異なります。表 1 は、場所により異なる Web サービスを呼び出すさまざまな方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="97675-p102">The way that you call a web service varies based on where the web service is located. Table 1 lists the different ways that you can call a web service based on location.</span></span>


<span data-ttu-id="97675-108">**表 1.Outlook アドインから web サービスを呼び出す方法**</span><span class="sxs-lookup"><span data-stu-id="97675-108">**Table 1. Ways to call web services from an Outlook add-in**</span></span>

<br/>

|<span data-ttu-id="97675-109">**Web サービスの場所**</span><span class="sxs-lookup"><span data-stu-id="97675-109">**Web service location**</span></span>|<span data-ttu-id="97675-110">**Web サービスを呼び出す方法**</span><span class="sxs-lookup"><span data-stu-id="97675-110">**Way to call the web service**</span></span>|
|:-----|:-----|
|<span data-ttu-id="97675-111">クライアント メールボックスをホストする Exchange サーバー</span><span class="sxs-lookup"><span data-stu-id="97675-111">The Exchange server that hosts the client mailbox</span></span>|<span data-ttu-id="97675-p103">[mailbox.makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) メソッドを使用して、アドインがサポートしている EWS 操作を呼び出します。また、メールボックスをホストしている Exchange サーバーも EWS を公開します。</span><span class="sxs-lookup"><span data-stu-id="97675-p103">Use the [mailbox.makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) method to call EWS operations that add-ins support. The Exchange server that hosts the mailbox also exposes EWS.</span></span>|
|<span data-ttu-id="97675-114">アドインの UI のソースの場所を提供する Web サーバー</span><span class="sxs-lookup"><span data-stu-id="97675-114">The web server that provides the source location for the add-in UI</span></span>|<span data-ttu-id="97675-p104">標準の JavaScript の手法を使用して Web サービスを呼び出します。UI フレーム内の JavaScript コードは、UI を提供する Web サーバーのコンテキストで実行されます。そのため、クロスサイト スクリプト エラーを発生させることなく、そのサーバーで Web サービスを呼び出すことができます。</span><span class="sxs-lookup"><span data-stu-id="97675-p104">Call the web service by using standard JavaScript techniques. The JavaScript code in the UI frame runs in the context of the web server that provides the UI. Therefore, it can call web services on that server without causing a cross-site scripting error.</span></span>|
|<span data-ttu-id="97675-118">上記以外の場所</span><span class="sxs-lookup"><span data-stu-id="97675-118">All other locations</span></span>|<span data-ttu-id="97675-p105">UI のソースの場所を提供する Web サーバー上で、Web サービスのプロキシを作成します。プロキシを作成しないと、クロスサイト スクリプト エラーによってアドインを実行できなくなります。プロキシを作成する方法の 1 つは JSON/P を使用することです。詳細については、「 [Office アドインのプライバシーとセキュリティ](../concepts/privacy-and-security.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="97675-p105">Create a proxy for the web service on the web server that provides the source location for the UI. If you do not provide a proxy, cross-site scripting errors will prevent your add-in from running. One way to provide a proxy is by using JSON/P. For more information, see [Privacy and security for Office Add-ins](../concepts/privacy-and-security.md).</span></span>|

## <a name="using-the-makeewsrequestasync-method-to-access-ews-operations"></a><span data-ttu-id="97675-123">makeEwsRequestAsync メソッドを使用して EWS 操作にアクセスする</span><span class="sxs-lookup"><span data-stu-id="97675-123">Using the makeEwsRequestAsync method to access EWS operations</span></span>

<span data-ttu-id="97675-124">[mailbox.makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) メソッドを使用して、ユーザーのメールボックスをホストしている Exchange サーバーに EWS 要求を行うことができます。</span><span class="sxs-lookup"><span data-stu-id="97675-124">You can use the [mailbox.makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) method to make an EWS request to the Exchange server that hosts the user's mailbox.</span></span>

<span data-ttu-id="97675-p106">EWS は、Exchange サーバーでの種々の操作をサポートしています。たとえば、アイテム レベルの操作としては、アイテムのコピー、検索、更新、または送信などがあり、フォルダー レベルの操作としては、フォルダーの作成、取得、または更新などがあります。EWS 操作を実行するには、その操作の XML SOAP 要求を作成します。操作が終了すると、その操作に関係するデータが含まれた XML SOAP 応答を受信します。EWS SOAP の要求と応答は、Messages.xsd ファイルで定義されているスキーマに従います。Messages.xsd ファイルは、他の EWS スキーマ ファイルと同様、EWS をホストしている IIS 仮想ディレクトリに配置されています。</span><span class="sxs-lookup"><span data-stu-id="97675-p106">EWS supports different operations on an Exchange server; for example, item-level operations to copy, find, update, or send an item, and folder-level operations to create, get, or update a folder. To perform an EWS operation, create an XML SOAP request for that operation. When the operation finishes, you get an XML SOAP response that contains data that is relevant to the operation. EWS SOAP requests and responses follow the schema defined in the Messages.xsd file. Like other EWS schema files, the Message.xsd file is located in the IIS virtual directory that hosts EWS.</span></span>

<span data-ttu-id="97675-130">メソッドを使用して EWS 操作を開始するには、 `makeEwsRequestAsync` 次の情報を指定します。</span><span class="sxs-lookup"><span data-stu-id="97675-130">To use the `makeEwsRequestAsync` method to initiate an EWS operation, provide the following:</span></span>

- <span data-ttu-id="97675-131">その EWS 操作の SOAP 要求に対する XML ( _data_ パラメーターへの引数)</span><span class="sxs-lookup"><span data-stu-id="97675-131">The XML for the SOAP request for that EWS operation, as an argument to the  _data_ parameter</span></span>

- <span data-ttu-id="97675-132">コールバック メソッド ( _callback_ 引数)</span><span class="sxs-lookup"><span data-stu-id="97675-132">A callback method (as the  _callback_ argument)</span></span>

- <span data-ttu-id="97675-133">そのコールバック メソッドに対するオプションの入力データ ( _userContext_ 引数)</span><span class="sxs-lookup"><span data-stu-id="97675-133">Any optional input data for that callback method (as the  _userContext_ argument)</span></span>

<span data-ttu-id="97675-p107">EWS SOAP 要求が完了すると、Outlook はコールバックメソッドを1つの引数 ( [AsyncResult](/javascript/api/office/office.asyncresult) オブジェクト) で呼び出します。コールバックメソッドは、オブジェクトの2つのプロパティにアクセスでき `AsyncResult` ます。この `value` プロパティには、EWS 操作の XML SOAP 応答と、必要に応じて、 `asyncContext` パラメーターとして渡されるデータが格納されているプロパティが含まれます。 `userContext` 通常、コールバックメソッドは SOAP 応答内の XML を解析して、関連する情報を取得し、その情報を適切に処理します。</span><span class="sxs-lookup"><span data-stu-id="97675-p107">When the EWS SOAP request is complete, Outlook calls the callback method with one argument, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object. The callback method can access two properties of the `AsyncResult` object: the `value` property, which contains the XML SOAP response of the EWS operation, and optionally, the `asyncContext` property, which contains any data passed as the `userContext` parameter. Typically, the callback method then parses the XML in the SOAP response to get any relevant information, and processes that information accordingly.</span></span>


## <a name="tips-for-parsing-ews-responses"></a><span data-ttu-id="97675-137">EWS 応答を解析するためのヒント</span><span class="sxs-lookup"><span data-stu-id="97675-137">Tips for parsing EWS responses</span></span>

<span data-ttu-id="97675-138">SOAP 応答を EWS 操作から解析する場合、ブラウザーに依存する以下の問題に注意してください。</span><span class="sxs-lookup"><span data-stu-id="97675-138">When parsing a SOAP response from an EWS operation, note the following browser-dependent issues:</span></span>


- <span data-ttu-id="97675-139">DOM メソッドを使用する場合は、タグ名に接頭辞を指定して、 `getElementsByTagName` Internet Explorer のサポートを組み込むことができます。</span><span class="sxs-lookup"><span data-stu-id="97675-139">Specify the prefix for a tag name when using the DOM method `getElementsByTagName`, to include support for Internet Explorer.</span></span>

  <span data-ttu-id="97675-p108">`getElementsByTagName` ブラウザーの種類によって動作が異なります。たとえば、EWS 応答には次の XML を含めることができます (表示のために書式指定され、省略されています)。</span><span class="sxs-lookup"><span data-stu-id="97675-p108">`getElementsByTagName` behaves differently depending on browser type. For example, an EWS response can contain the following XML (formatted and abbreviated for display purposes):</span></span>

   ```XML
        <t:ExtendedProperty><t:ExtendedFieldURI PropertySetId="00000000-0000-0000-0000-000000000000" 
        PropertyName="MyProperty" 
        PropertyType="String"/>
        <t:Value>{
        ...
        }</t:Value></t:ExtendedProperty>
   ```

   <span data-ttu-id="97675-142">次のようにコードを使用すると、Chrome などのブラウザーでタグで囲まれた XML を取得でき `ExtendedProperty` ます。</span><span class="sxs-lookup"><span data-stu-id="97675-142">Code, as in the following, would work on a browser like Chrome to get the XML enclosed by the `ExtendedProperty` tags:</span></span>

   ```js
        var mailbox = Office.context.mailbox;
        mailbox.makeEwsRequestAsync(mailbox.item.itemId, function(result) {
            var response = $.parseXML(result.value);
            var extendedProps = response.getElementsByTagName("ExtendedProperty")
            });
   ```

   <span data-ttu-id="97675-143">Internet Explorer では、以下に示すように、タグ名に接頭辞 `t:` を含める必要があります。</span><span class="sxs-lookup"><span data-stu-id="97675-143">On Internet Explorer, you must include the `t:` prefix of the tag name, as shown below:</span></span>

   ```js
        var mailbox = Office.context.mailbox;
        mailbox.makeEwsRequestAsync(mailbox.item.itemId, function(result) {
            var response = $.parseXML(result.value);
            var extendedProps = response.getElementsByTagName("t:ExtendedProperty")
            });
   ```

- <span data-ttu-id="97675-144">次に示すように、DOM プロパティを使用し `textContent` て、EWS 応答のタグの内容を取得します。</span><span class="sxs-lookup"><span data-stu-id="97675-144">Use the DOM property `textContent` to get the contents of a tag in an EWS response, as shown below:</span></span>

   ```js
      content = $.parseJSON(value.textContent);
   ```

   <span data-ttu-id="97675-145">`innerHTML`Internet Explorer では、などの他のプロパティは EWS 応答のいくつかのタグに対して機能しない場合があります。</span><span class="sxs-lookup"><span data-stu-id="97675-145">Other properties such as `innerHTML` may not work on Internet Explorer for some tags in an EWS response.</span></span>


## <a name="example"></a><span data-ttu-id="97675-146">例</span><span class="sxs-lookup"><span data-stu-id="97675-146">Example</span></span>

<span data-ttu-id="97675-p109">次の例は、 `makeEwsRequestAsync` [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) 操作を使用してアイテムの件名を取得するために呼び出しを行います。この例には、次の3つの関数が含まれています。</span><span class="sxs-lookup"><span data-stu-id="97675-p109">The following example calls `makeEwsRequestAsync` to use the [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to get the subject of an item. This example includes the following three functions:</span></span>

-  <span data-ttu-id="97675-149">`getSubjectRequest`&ndash;入力としてアイテム ID を受け取り、 `GetItem` 指定されたアイテムに対して呼び出す SOAP 要求の XML を返します。</span><span class="sxs-lookup"><span data-stu-id="97675-149">`getSubjectRequest` &ndash; Takes an item ID as input, and returns the XML for the SOAP request to call `GetItem` for the specified item.</span></span>

-  <span data-ttu-id="97675-150">`sendRequest`&ndash; `getSubjectRequest` [呼び出し] 選択したアイテムの soap 要求を取得し、soap 要求とコールバックメソッドを渡して、 `callback` `makeEwsRequestAsync` 指定されたアイテムの件名を取得します。</span><span class="sxs-lookup"><span data-stu-id="97675-150">`sendRequest` &ndash; Calls  `getSubjectRequest` to get the SOAP request for the selected item, then passes the SOAP request and the callback method, `callback`, to `makeEwsRequestAsync` to get the subject of the specified item.</span></span>

-  <span data-ttu-id="97675-151">`callback` &ndash; 指定のアイテムの件名とその他の情報が含まれている SOAP 応答を処理します。</span><span class="sxs-lookup"><span data-stu-id="97675-151">`callback` &ndash; Processes the SOAP response which includes any subject and other information about the specified item.</span></span>


```js
function getSubjectRequest(id) {
   // Return a GetItem operation request for the subject of the specified item. 
   var result = 
'<?xml version="1.0" encoding="utf-8"?>' +
'<soap:Envelope xmlns:xsi="https://www.w3.org/2001/XMLSchema-instance"' +
'               xmlns:xsd="https://www.w3.org/2001/XMLSchema"' +
'               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"' +
'               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
'  <soap:Header>' +
'    <RequestServerVersion Version="Exchange2013" xmlns="http://schemas.microsoft.com/exchange/services/2006/types" soap:mustUnderstand="0" />' +
'  </soap:Header>' +
'  <soap:Body>' +
'    <GetItem xmlns="http://schemas.microsoft.com/exchange/services/2006/messages">' +
'      <ItemShape>' +
'        <t:BaseShape>IdOnly</t:BaseShape>' +
'        <t:AdditionalProperties>' +
'            <t:FieldURI FieldURI="item:Subject"/>' +
'        </t:AdditionalProperties>' +
'      </ItemShape>' +
'      <ItemIds><t:ItemId Id="' + id + '"/></ItemIds>' +
'    </GetItem>' +
'  </soap:Body>' +
'</soap:Envelope>';

   return result;
}

function sendRequest() {
   // Create a local variable that contains the mailbox.
   var mailbox = Office.context.mailbox;

   mailbox.makeEwsRequestAsync(getSubjectRequest(mailbox.item.itemId), callback);
}

function callback(asyncResult)  {
   var result = asyncResult.value;
   var context = asyncResult.context;

   // Process the returned response here.
}
```


## <a name="ews-operations-that-add-ins-support"></a><span data-ttu-id="97675-152">アドインでサポートしている EWS 操作</span><span class="sxs-lookup"><span data-stu-id="97675-152">EWS operations that add-ins support</span></span>

<span data-ttu-id="97675-p110">Outlook アドインは、メソッドによって EWS で利用可能な操作のサブセットにアクセスでき `makeEwsRequestAsync` ます。EWS 操作についてよく知らない場合、またはメソッドを使用して操作にアクセスする方法については、最初に SOAP 要求の例を使用して、 `makeEwsRequestAsync` _データ_ 引数をカスタマイズしてください。</span><span class="sxs-lookup"><span data-stu-id="97675-p110">Outlook add-ins can access a subset of operations that are available in EWS via the `makeEwsRequestAsync` method. If you are unfamiliar with EWS operations and how to use the `makeEwsRequestAsync` method to access an operation, start with a SOAP request example to customize your _data_ argument.</span></span>

<span data-ttu-id="97675-155">次に、メソッドの使用方法を説明し `makeEwsRequestAsync` ます。</span><span class="sxs-lookup"><span data-stu-id="97675-155">The following describes how you can use the `makeEwsRequestAsync` method:</span></span>

1. <span data-ttu-id="97675-156">XML 内のアイテム ID および関係する EWS 操作属性を適切な値に置き換えます。</span><span class="sxs-lookup"><span data-stu-id="97675-156">In the XML, substitute any item IDs and relevant EWS operation attributes with appropriate values.</span></span>

2. <span data-ttu-id="97675-157">SOAP 要求をの  _data_ パラメーターの引数として含め `makeEwsRequestAsync` ます。</span><span class="sxs-lookup"><span data-stu-id="97675-157">Include the SOAP request as an argument for the  _data_ parameter of `makeEwsRequestAsync`.</span></span>

3. <span data-ttu-id="97675-158">コールバックメソッドを指定して呼び出しを `makeEwsRequestAsync` 行います。</span><span class="sxs-lookup"><span data-stu-id="97675-158">Specify a callback method and call `makeEwsRequestAsync`.</span></span>

4. <span data-ttu-id="97675-159">コールバック メソッド内で、SOAP 応答内の操作の結果を検証します。</span><span class="sxs-lookup"><span data-stu-id="97675-159">In the callback method, verify the results of the operation in the SOAP response.</span></span>

5. <span data-ttu-id="97675-160">必要に応じて EWS 操作の結果を使用します。</span><span class="sxs-lookup"><span data-stu-id="97675-160">Use the results of the EWS operation according to your needs.</span></span>

<span data-ttu-id="97675-p111">次の表は、アドインがサポートしている EWS 操作を示しています。SOAP の要求と応答の例を表示するには、各操作のリンクを選択します。EWS 操作の詳細については、「 [Exchange での EWS の操作](/exchange/client-developer/web-service-reference/ews-operations-in-exchange)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="97675-p111">The following table lists the EWS operations that add-ins support. To see examples of SOAP requests and responses, choose the link for each operation. For more information about EWS operations, see [EWS operations in Exchange](/exchange/client-developer/web-service-reference/ews-operations-in-exchange).</span></span>

<span data-ttu-id="97675-164">**表 2サポートされている EWS 操作**</span><span class="sxs-lookup"><span data-stu-id="97675-164">**Table 2. Supported EWS operations**</span></span>

<br/>

|<span data-ttu-id="97675-165">**EWS 操作**</span><span class="sxs-lookup"><span data-stu-id="97675-165">**EWS operation**</span></span>|<span data-ttu-id="97675-166">**説明**</span><span class="sxs-lookup"><span data-stu-id="97675-166">**Description**</span></span>|
|:-----|:-----|
|<span data-ttu-id="97675-167">
  [CopyItem 操作](/exchange/client-developer/web-service-reference/copyitem-operation)</span><span class="sxs-lookup"><span data-stu-id="97675-167">[CopyItem operation](/exchange/client-developer/web-service-reference/copyitem-operation)</span></span>|<span data-ttu-id="97675-168">指定したアイテムをコピーし、Exchange ストア内の指定のフォルダーに新しいアイテムを入れます。</span><span class="sxs-lookup"><span data-stu-id="97675-168">Copies the specified items and puts the new items in a designated folder in the Exchange store.</span></span>|
|<span data-ttu-id="97675-169">
  [CreateFolder 操作](/exchange/client-developer/web-service-reference/createfolder-operation)</span><span class="sxs-lookup"><span data-stu-id="97675-169">[CreateFolder operation](/exchange/client-developer/web-service-reference/createfolder-operation)</span></span>|<span data-ttu-id="97675-170">Exchange ストア内の指定の場所にフォルダーを作成します。</span><span class="sxs-lookup"><span data-stu-id="97675-170">Creates folders in the specified location in the Exchange store.</span></span>|
|<span data-ttu-id="97675-171">
  [CreateItem 操作](/exchange/client-developer/web-service-reference/createitem-operation)</span><span class="sxs-lookup"><span data-stu-id="97675-171">[CreateItem operation](/exchange/client-developer/web-service-reference/createitem-operation)</span></span>|<span data-ttu-id="97675-172">Exchange ストアに指定したアイテムを作成します。</span><span class="sxs-lookup"><span data-stu-id="97675-172">Creates the specified items in the Exchange store.</span></span>|
|[<span data-ttu-id="97675-173">ExpandDL 操作</span><span class="sxs-lookup"><span data-stu-id="97675-173">ExpandDL operation</span></span>](/exchange/client-developer/web-service-reference/expanddl-operation)|<span data-ttu-id="97675-174">配布リストの完全なメンバシップを表示します。</span><span class="sxs-lookup"><span data-stu-id="97675-174">Displays the full membership of distribution lists.</span></span>|
|[<span data-ttu-id="97675-175">FindConversation 操作</span><span class="sxs-lookup"><span data-stu-id="97675-175">FindConversation operation</span></span>](/exchange/client-developer/web-service-reference/findconversation-operation)|<span data-ttu-id="97675-176">Exchange ストアの指定したフォルダー内のスレッドのリストを列挙します。</span><span class="sxs-lookup"><span data-stu-id="97675-176">Enumerates a list of conversations in the specified folder in the Exchange store.</span></span>|
|<span data-ttu-id="97675-177">
  [FindFolder 操作](/exchange/client-developer/web-service-reference/findfolder-operation)</span><span class="sxs-lookup"><span data-stu-id="97675-177">[FindFolder operation](/exchange/client-developer/web-service-reference/findfolder-operation)</span></span>|<span data-ttu-id="97675-178">指定したフォルダーのサブフォルダーを検索し、一群のサブフォルダーを記述した一群のプロパティを返します。</span><span class="sxs-lookup"><span data-stu-id="97675-178">Finds subfolders of an identified folder and returns a set of properties that describe the set of subfolders.</span></span>|
|<span data-ttu-id="97675-179">
  [FindItem 操作](/exchange/client-developer/web-service-reference/finditem-operation)</span><span class="sxs-lookup"><span data-stu-id="97675-179">[FindItem operation](/exchange/client-developer/web-service-reference/finditem-operation)</span></span>|<span data-ttu-id="97675-180">Exchange ストアの指定したフォルダー内にあるアイテムを特定します。</span><span class="sxs-lookup"><span data-stu-id="97675-180">Identifies items that are located in a specified folder in the Exchange store.</span></span>|
|<span data-ttu-id="97675-181">
  [GetConversationItems 操作](/exchange/client-developer/web-service-reference/getconversationitems-operation)</span><span class="sxs-lookup"><span data-stu-id="97675-181">[GetConversationItems operation](/exchange/client-developer/web-service-reference/getconversationitems-operation)</span></span>|<span data-ttu-id="97675-182">スレッドでノードに整理された 1 つまたは複数のアイテム セットを取得します。</span><span class="sxs-lookup"><span data-stu-id="97675-182">Gets one or more sets of items that are organized in nodes in a conversation.</span></span>|
|<span data-ttu-id="97675-183">
  [GetFolder 操作](/exchange/client-developer/web-service-reference/getfolder-operation)</span><span class="sxs-lookup"><span data-stu-id="97675-183">[GetFolder operation](/exchange/client-developer/web-service-reference/getfolder-operation)</span></span>|<span data-ttu-id="97675-184">Exchange ストアから、指定したプロパティとフォルダーの内容を取得します。</span><span class="sxs-lookup"><span data-stu-id="97675-184">Gets the specified properties and contents of folders from the Exchange store.</span></span>|
|<span data-ttu-id="97675-185">
  [GetItem 操作](/exchange/client-developer/web-service-reference/getitem-operation)</span><span class="sxs-lookup"><span data-stu-id="97675-185">[GetItem operation](/exchange/client-developer/web-service-reference/getitem-operation)</span></span>|<span data-ttu-id="97675-186">Exchange ストアから指定したプロパティとアイテムの内容を取得します。</span><span class="sxs-lookup"><span data-stu-id="97675-186">Gets the specified properties and contents of items from the Exchange store.</span></span>|
|[<span data-ttu-id="97675-187">GetUserAvailability 操作</span><span class="sxs-lookup"><span data-stu-id="97675-187">GetUserAvailability operation</span></span>](/exchange/client-developer/web-service-reference/getuseravailability-operation)|<span data-ttu-id="97675-188">指定された期間のユーザー、部屋、およびリソースのセットの詳細な空き時間情報を提供します。</span><span class="sxs-lookup"><span data-stu-id="97675-188">Provides detailed information about the availability of a set of users, rooms, and resources within a specified time period.</span></span>|
|[<span data-ttu-id="97675-189">MarkAsJunk 操作</span><span class="sxs-lookup"><span data-stu-id="97675-189">MarkAsJunk operation</span></span>](/exchange/client-developer/web-service-reference/markasjunk-operation)|<span data-ttu-id="97675-190">電子メール メッセージを [迷惑メール] フォルダーに移動し、それらのメッセージの差出人を受信拒否リストに追加したり、リストから削除したりします。</span><span class="sxs-lookup"><span data-stu-id="97675-190">Moves email messages to the Junk Email folder, and adds or removes senders of the messages from the blocked senders list accordingly.</span></span>|
|<span data-ttu-id="97675-191">
  [MoveItem 操作](/exchange/client-developer/web-service-reference/moveitem-operation)</span><span class="sxs-lookup"><span data-stu-id="97675-191">[MoveItem operation](/exchange/client-developer/web-service-reference/moveitem-operation)</span></span>|<span data-ttu-id="97675-192">アイテムを Exchange ストア内の単一のフォルダーに移動します。</span><span class="sxs-lookup"><span data-stu-id="97675-192">Moves items to a single destination folder in the Exchange store.</span></span>|
|[<span data-ttu-id="97675-193">ResolveNames 操作</span><span class="sxs-lookup"><span data-stu-id="97675-193">ResolveNames operation</span></span>](/exchange/client-developer/web-service-reference/resolvenames-operation)|<span data-ttu-id="97675-194">あいまいな電子メール アドレスおよび表示名を解決します。</span><span class="sxs-lookup"><span data-stu-id="97675-194">Resolves ambiguous email addresses and display names.</span></span>|
|[<span data-ttu-id="97675-195">SendItem 操作</span><span class="sxs-lookup"><span data-stu-id="97675-195">SendItem operation</span></span>](/exchange/client-developer/web-service-reference/senditem-operation)|<span data-ttu-id="97675-196">Exchange ストアにある電子メール メッセージを送信します。</span><span class="sxs-lookup"><span data-stu-id="97675-196">Sends email messages that are located in the Exchange store.</span></span>|
|<span data-ttu-id="97675-197">
  [UpdateFolder 操作](/exchange/client-developer/web-service-reference/updatefolder-operation)</span><span class="sxs-lookup"><span data-stu-id="97675-197">[UpdateFolder operation](/exchange/client-developer/web-service-reference/updatefolder-operation)</span></span>|<span data-ttu-id="97675-198">Exchange ストアの既存のフォルダーのプロパティを変更します。</span><span class="sxs-lookup"><span data-stu-id="97675-198">Modifies the properties of existing folders in the Exchange store.</span></span>|
|<span data-ttu-id="97675-199">
  [UpdateItem 操作](/exchange/client-developer/web-service-reference/updateitem-operation)</span><span class="sxs-lookup"><span data-stu-id="97675-199">[UpdateItem operation](/exchange/client-developer/web-service-reference/updateitem-operation)</span></span>|<span data-ttu-id="97675-200">Exchange ストアの既存のアイテムのプロパティを変更します。</span><span class="sxs-lookup"><span data-stu-id="97675-200">Modifies the properties of existing items in the Exchange store.</span></span>|

 > [!NOTE]
 > <span data-ttu-id="97675-201">FAI (フォルダー関連情報) アイテムをアドインから更新 (または作成) することはできません。</span><span class="sxs-lookup"><span data-stu-id="97675-201">FAI (Folder Associated Information) items cannot be updated (or created) from an add-in.</span></span> <span data-ttu-id="97675-202">これらの非表示メッセージはフォルダーに保存され、さまざまな設定と補助データを格納するときに使用されます。</span><span class="sxs-lookup"><span data-stu-id="97675-202">These hidden messages are stored in a folder and are used to store a variety of settings and auxiliary data.</span></span>  <span data-ttu-id="97675-203">UpdateItem 操作を使用しようとすると、ErrorAccessDenied エラー (「Office 拡張機能はこのような種類のアイテムの更新を許可されていません」) がスローされます。</span><span class="sxs-lookup"><span data-stu-id="97675-203">Attempting to use the UpdateItem operation will throw an ErrorAccessDenied error: "Office extension is not allowed to update this type of item".</span></span> <span data-ttu-id="97675-204">代わりに、[EWS マネージ API](/exchange/client-developer/exchange-web-services/get-started-with-ews-managed-api-client-applications) を使用して、Windows クライアントまたはサーバー アプリケーションからこれらのアイテムを更新することができます。</span><span class="sxs-lookup"><span data-stu-id="97675-204">As an alternative, you may use the [EWS Managed API](/exchange/client-developer/exchange-web-services/get-started-with-ews-managed-api-client-applications) to update these items from a Windows client or a server application.</span></span> <span data-ttu-id="97675-205">内部の service-type データ構造は変更対象であり、ソリューションが破損する可能性があるため注意してください。</span><span class="sxs-lookup"><span data-stu-id="97675-205">Caution is recommended as internal, service-type data structures are subject to change and could break your solution.</span></span>


## <a name="authentication-and-permission-considerations-for-makeewsrequestasync"></a><span data-ttu-id="97675-206">makeEwsRequestAsync の認証とアクセス許可について</span><span class="sxs-lookup"><span data-stu-id="97675-206">Authentication and permission considerations for makeEwsRequestAsync</span></span>

<span data-ttu-id="97675-207">メソッドを使用する場合 `makeEwsRequestAsync` 、要求は現在のユーザーの電子メールアカウントの資格情報を使用して認証されます。</span><span class="sxs-lookup"><span data-stu-id="97675-207">When you use the `makeEwsRequestAsync` method, the request is authenticated by using the email account credentials of the current user.</span></span> <span data-ttu-id="97675-208">この `makeEwsRequestAsync` メソッドは、ユーザーの資格情報を管理して、要求に対して認証の資格情報を提供する必要がないようにします。</span><span class="sxs-lookup"><span data-stu-id="97675-208">The `makeEwsRequestAsync` method manages the credentials for you so that you do not have to provide authentication credentials with your request.</span></span>

> [!NOTE]
> <span data-ttu-id="97675-209">サーバー管理者は、 [set-webservicesvirtualdirectory](/powershell/module/exchange/client-access-servers/New-WebServicesVirtualDirectory?view=exchange-ps) または [set-webservicesvirtualdirectory](/powershell/module/exchange/client-access-servers/Set-WebServicesVirtualDirectory?view=exchange-ps) コマンドレットを使用して、クライアントアクセスサーバーの Ews ディレクトリで _oauthauthentication_ パラメーターを **TRUE** に設定し、 `makeEwsRequestAsync` メソッドで ews 要求を行うことができるようにする必要があります。</span><span class="sxs-lookup"><span data-stu-id="97675-209">The server administrator must use the [New-WebServicesVirtualDirectory](/powershell/module/exchange/client-access-servers/New-WebServicesVirtualDirectory?view=exchange-ps) or the [Set-WebServicesVirtualDirectory](/powershell/module/exchange/client-access-servers/Set-WebServicesVirtualDirectory?view=exchange-ps) cmdlet to set the _OAuthAuthentication_ parameter to **true** on the Client Access server EWS directory in order to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

<span data-ttu-id="97675-210">アドインは、 `ReadWriteMailbox` アドインマニフェストでメソッドを使用するためのアクセス許可を指定する必要があり `makeEwsRequestAsync` ます。</span><span class="sxs-lookup"><span data-stu-id="97675-210">Your add-in must specify the `ReadWriteMailbox` permission in its add-in manifest to use the `makeEwsRequestAsync` method.</span></span> <span data-ttu-id="97675-211">アクセス許可の使用の詳細については `ReadWriteMailbox` 、「 [ReadWriteMailbox Permission](understanding-outlook-add-in-permissions.md#readwritemailbox-permission) For [Outlook アドインのアクセス許可](understanding-outlook-add-in-permissions.md)について」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="97675-211">For information about using the `ReadWriteMailbox` permission, see the section [ReadWriteMailbox permission](understanding-outlook-add-in-permissions.md#readwritemailbox-permission) in [Understanding Outlook add-in permissions](understanding-outlook-add-in-permissions.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="97675-212">関連項目</span><span class="sxs-lookup"><span data-stu-id="97675-212">See also</span></span>

- [<span data-ttu-id="97675-213">Office アドインのプライバシーとセキュリティ</span><span class="sxs-lookup"><span data-stu-id="97675-213">Privacy and security for Office Add-ins</span></span>](../concepts/privacy-and-security.md)
- [<span data-ttu-id="97675-214">Office アドインにおける同一生成元ポリシーの制限への対処</span><span class="sxs-lookup"><span data-stu-id="97675-214">Addressing same-origin policy limitations in Office Add-ins</span></span>](../develop/addressing-same-origin-policy-limitations.md)
- <span data-ttu-id="97675-215">
  [Exchange 用 EWS リファレンス](/exchange/client-developer/web-service-reference/ews-reference-for-exchange)</span><span class="sxs-lookup"><span data-stu-id="97675-215">[EWS reference for Exchange](/exchange/client-developer/web-service-reference/ews-reference-for-exchange)</span></span>
- <span data-ttu-id="97675-216">
  [Exchange での Outlook 用メール アプリと EWS](/exchange/client-developer/exchange-web-services/mail-apps-for-outlook-and-ews-in-exchange)</span><span class="sxs-lookup"><span data-stu-id="97675-216">[Mail apps for Outlook and EWS in Exchange](/exchange/client-developer/exchange-web-services/mail-apps-for-outlook-and-ews-in-exchange)</span></span>

<span data-ttu-id="97675-217">ASP.NET Web API を使用してアドイン用のバックエンド サービスを作成する場合は、以下の資料を参照してください。</span><span class="sxs-lookup"><span data-stu-id="97675-217">See the following for creating backend services for add-ins using ASP.NET Web API:</span></span>

- [<span data-ttu-id="97675-218">ASP.NET Web API を使用して Office アドイン用 Web サービスを作成する</span><span class="sxs-lookup"><span data-stu-id="97675-218">Create a web service for an Office Add-in using the ASP.NET Web API</span></span>](https://blogs.msdn.microsoft.com/officeapps/2013/06/10/create-a-web-service-for-an-app-for-office-using-the-asp-net-web-api/)
- [<span data-ttu-id="97675-219">ASP.NET Web API を使用した HTTP サービスの構築に関する基本</span><span class="sxs-lookup"><span data-stu-id="97675-219">The basics of building an HTTP service using ASP.NET Web API</span></span>](https://www.asp.net/web-api)
