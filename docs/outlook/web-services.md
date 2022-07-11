---
title: Outlook アドインから Exchange Web サービス (EWS) を使用する
description: Outlook アドインが Exchange Web サービスに情報を要求する方法の例を示します。
ms.date: 07/08/2022
ms.localizationpriority: medium
ms.openlocfilehash: 6590967ef79e03cdbeee612199aba7a681b6dcdb
ms.sourcegitcommit: d8ea4b761f44d3227b7f2c73e52f0d2233bf22e2
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/11/2022
ms.locfileid: "66713064"
---
# <a name="call-web-services-from-an-outlook-add-in"></a>Outlook アドインから Web サービスを呼び出す

アドインは、Exchange Server 2013 が実行されているコンピューター、アドインの UI のソースの場所を提供するサーバー上で利用できる Web サービス、またはインターネット上で利用できる Web サービスから Exchange Web サービス (EWS) を使用できます。この記事では、Outlook アドインからどのように EWS の情報を要求できるかを示す例を説明します。

Web サービスを呼び出す方法は、Web サービスは配置された場所によって異なります。表 1 は、場所により異なる Web サービスを呼び出すさまざまな方法を示しています。

**表 1.Outlook アドインから web サービスを呼び出す方法**

|**Web サービスの場所**|**Web サービスを呼び出す方法**|
|:-----|:-----|
|クライアント メールボックスをホストする Exchange サーバー|[mailbox.makeEwsRequestAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) メソッドを使用して、アドインがサポートしている EWS 操作を呼び出します。また、メールボックスをホストしている Exchange サーバーも EWS を公開します。|
|アドインの UI のソースの場所を提供する Web サーバー|標準の JavaScript の手法を使用して Web サービスを呼び出します。UI フレーム内の JavaScript コードは、UI を提供する Web サーバーのコンテキストで実行されます。そのため、クロスサイト スクリプト エラーを発生させることなく、そのサーバーで Web サービスを呼び出すことができます。|
|上記以外の場所|UI のソースの場所を提供する Web サーバー上で、Web サービスのプロキシを作成します。プロキシを作成しないと、クロスサイト スクリプト エラーによってアドインを実行できなくなります。プロキシを作成する方法の 1 つは JSON/P を使用することです。詳細については、「 [Office アドインのプライバシーとセキュリティ](../concepts/privacy-and-security.md)」を参照してください。|

## <a name="using-the-makeewsrequestasync-method-to-access-ews-operations"></a>makeEwsRequestAsync メソッドを使用して EWS 操作にアクセスする

[mailbox.makeEwsRequestAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) メソッドを使用して、ユーザーのメールボックスをホストしている Exchange サーバーに EWS 要求を行うことができます。

EWS は、Exchange サーバーでの種々の操作をサポートしています。たとえば、アイテム レベルの操作としては、アイテムのコピー、検索、更新、または送信などがあり、フォルダー レベルの操作としては、フォルダーの作成、取得、または更新などがあります。EWS 操作を実行するには、その操作の XML SOAP 要求を作成します。操作が終了すると、その操作に関係するデータが含まれた XML SOAP 応答を受信します。EWS SOAP の要求と応答は、Messages.xsd ファイルで定義されているスキーマに従います。Messages.xsd ファイルは、他の EWS スキーマ ファイルと同様、EWS をホストしている IIS 仮想ディレクトリに配置されています。

メソッドを `makeEwsRequestAsync` 使用して EWS 操作を開始するには、次を指定します。

- その EWS 操作の SOAP 要求に対する XML ( _data_ パラメーターへの引数)

- コールバック メソッド ( _callback_ 引数)

- そのコールバック メソッドに対するオプションの入力データ ( _userContext_ 引数)

EWS SOAP 要求が完了すると、Outlook は 1 つの引数 ( [AsyncResult](/javascript/api/office/office.asyncresult) オブジェクト) でコールバック メソッドを呼び出します。 コールバック メソッドは、EWS 操作の `AsyncResult` XML SOAP 応答を `value` 含むプロパティと、必要に応じてパラメーターとして渡されたデータを `asyncContext` 含むプロパティの 2 つのプロパティに `userContext` アクセスできます。 通常、コールバック メソッドは SOAP 応答内の XML を解析して関連情報を取得し、それに応じてその情報を処理します。

## <a name="tips-for-parsing-ews-responses"></a>EWS 応答を解析するためのヒント

EWS 操作から SOAP 応答を解析する場合は、ブラウザーに依存する次の問題に注意してください。

- INTERNET Explorer のサポートを含めるには、DOM メソッド `getElementsByTagName`を使用するときにタグ名のプレフィックスを指定します。

  `getElementsByTagName` ブラウザーの種類によって動作が異なります。 たとえば、EWS 応答には、次の XML を含めることができます (表示目的で書式設定され、省略されています)。

   ```XML
   <t:ExtendedProperty><t:ExtendedFieldURI PropertySetId="00000000-0000-0000-0000-000000000000" 
   PropertyName="MyProperty" 
   PropertyType="String"/>
   <t:Value>{
   ...
   }</t:Value></t:ExtendedProperty>
   ```

   次のように、コードは Chrome などのブラウザーで動作し、タグで囲まれた XML を `ExtendedProperty` 取得します。

   ```js
   const mailbox = Office.context.mailbox;
   mailbox.makeEwsRequestAsync(mailbox.item.itemId, function(result) {
       const response = $.parseXML(result.value);
       const extendedProps = response.getElementsByTagName("ExtendedProperty")
   });
   ```

   Internet Explorer では、次のようにタグ名のプレフィックスを含める `t:` 必要があります。

   ```js
   const mailbox = Office.context.mailbox;
   mailbox.makeEwsRequestAsync(mailbox.item.itemId, function(result) {
       const response = $.parseXML(result.value);
       const extendedProps = response.getElementsByTagName("t:ExtendedProperty")
   });
   ```

- 次のように、DOM プロパティ `textContent` を使用して EWS 応答内のタグの内容を取得します。

   ```js
   content = $.parseJSON(value.textContent);
   ```

   EWS 応答内の一部のタグについては、Internet Explorer では機能しないなどの `innerHTML` その他のプロパティがあります。

## <a name="example"></a>例

次の例では、[GetItem](/exchange/client-developer/web-service-reference/getitem-operation) 操作を使用して項目の件名を取得する呼び出し`makeEwsRequestAsync`の例を示します。 この例には、次の 3 つの関数が含まれています。

- `getSubjectRequest`&ndash;項目 ID を入力として受け取り、指定した項目を呼び出`GetItem`す SOAP 要求の XML を返します。

- `sendRequest`&ndash;選択した項目の SOAP 要求を取得し、SOAP 要求とコールバック メソッドを渡して、`callback`指定した項目のサブジェクトを取得する`makeEwsRequestAsync`呼び出`getSubjectRequest`し。

- `callback` &ndash; 指定のアイテムの件名とその他の情報が含まれている SOAP 応答を処理します。

```js
function getSubjectRequest(id) {
   // Return a GetItem operation request for the subject of the specified item. 
   const result = 
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
   const mailbox = Office.context.mailbox;

   mailbox.makeEwsRequestAsync(getSubjectRequest(mailbox.item.itemId), callback);
}

function callback(asyncResult)  {
   const result = asyncResult.value;
   const context = asyncResult.context;

   // Process the returned response here.
}
```

## <a name="ews-operations-that-add-ins-support"></a>アドインでサポートしている EWS 操作

Outlook アドインは、メソッドを使用して EWS で使用できる操作のサブセットに `makeEwsRequestAsync` アクセスできます。 EWS 操作に慣れていない場合や、メソッドを使用 `makeEwsRequestAsync` して操作にアクセスする方法については、SOAP 要求の例から始めて _データ_ 引数をカスタマイズします。

次に、メソッドを使用 `makeEwsRequestAsync` する方法について説明します。

1. XML 内のアイテム ID および関係する EWS 操作属性を適切な値に置き換えます。

1. の  _データ_ パラメーター `makeEwsRequestAsync`の引数として SOAP 要求を含めます。

1. コールバック メソッドを指定し、呼び出します `makeEwsRequestAsync`。

1. コールバック メソッド内で、SOAP 応答内の操作の結果を検証します。

1. 必要に応じて EWS 操作の結果を使用します。

次の表は、アドインがサポートしている EWS 操作を示しています。SOAP の要求と応答の例を表示するには、各操作のリンクを選択します。EWS 操作の詳細については、「 [Exchange での EWS の操作](/exchange/client-developer/web-service-reference/ews-operations-in-exchange)」を参照してください。

**表 2サポートされている EWS 操作**

|**EWS 操作**|**説明**|
|:-----|:-----|
|
  [CopyItem 操作](/exchange/client-developer/web-service-reference/copyitem-operation)|指定したアイテムをコピーし、Exchange ストア内の指定のフォルダーに新しいアイテムを入れます。|
|
  [CreateFolder 操作](/exchange/client-developer/web-service-reference/createfolder-operation)|Exchange ストア内の指定の場所にフォルダーを作成します。|
|
  [CreateItem 操作](/exchange/client-developer/web-service-reference/createitem-operation)|Exchange ストアに指定したアイテムを作成します。|
|[ExpandDL 操作](/exchange/client-developer/web-service-reference/expanddl-operation)|配布リストの完全なメンバシップを表示します。|
|[FindConversation 操作](/exchange/client-developer/web-service-reference/findconversation-operation)|Exchange ストアの指定したフォルダー内のスレッドのリストを列挙します。|
|
  [FindFolder 操作](/exchange/client-developer/web-service-reference/findfolder-operation)|指定したフォルダーのサブフォルダーを検索し、一群のサブフォルダーを記述した一群のプロパティを返します。|
|
  [FindItem 操作](/exchange/client-developer/web-service-reference/finditem-operation)|Exchange ストアの指定したフォルダー内にあるアイテムを特定します。|
|
  [GetConversationItems 操作](/exchange/client-developer/web-service-reference/getconversationitems-operation)|スレッドでノードに整理された 1 つまたは複数のアイテム セットを取得します。|
|
  [GetFolder 操作](/exchange/client-developer/web-service-reference/getfolder-operation)|Exchange ストアから、指定したプロパティとフォルダーの内容を取得します。|
|
  [GetItem 操作](/exchange/client-developer/web-service-reference/getitem-operation)|Exchange ストアから指定したプロパティとアイテムの内容を取得します。|
|[GetUserAvailability 操作](/exchange/client-developer/web-service-reference/getuseravailability-operation)|指定された期間のユーザー、部屋、およびリソースのセットの詳細な空き時間情報を提供します。|
|[MarkAsJunk 操作](/exchange/client-developer/web-service-reference/markasjunk-operation)|電子メール メッセージを [迷惑メール] フォルダーに移動し、それらのメッセージの差出人を受信拒否リストに追加したり、リストから削除したりします。|
|
  [MoveItem 操作](/exchange/client-developer/web-service-reference/moveitem-operation)|アイテムを Exchange ストア内の単一のフォルダーに移動します。|
|[ResolveNames 操作](/exchange/client-developer/web-service-reference/resolvenames-operation)|あいまいな電子メール アドレスおよび表示名を解決します。|
|[SendItem 操作](/exchange/client-developer/web-service-reference/senditem-operation)|Exchange ストアにある電子メール メッセージを送信します。|
|
  [UpdateFolder 操作](/exchange/client-developer/web-service-reference/updatefolder-operation)|Exchange ストアの既存のフォルダーのプロパティを変更します。|
|
  [UpdateItem 操作](/exchange/client-developer/web-service-reference/updateitem-operation)|Exchange ストアの既存のアイテムのプロパティを変更します。|

 > [!NOTE]
 > FAI (フォルダー関連情報) アイテムをアドインから更新 (または作成) することはできません。 これらの非表示メッセージはフォルダーに保存され、さまざまな設定と補助データを格納するときに使用されます。  UpdateItem 操作を使用しようとすると、ErrorAccessDenied エラー (「Office 拡張機能はこのような種類のアイテムの更新を許可されていません」) がスローされます。 代わりに、[EWS マネージ API](/exchange/client-developer/exchange-web-services/get-started-with-ews-managed-api-client-applications) を使用して、Windows クライアントまたはサーバー アプリケーションからこれらのアイテムを更新することができます。 内部の service-type データ構造は変更対象であり、ソリューションが破損する可能性があるため注意してください。

## <a name="authentication-and-permission-considerations-for-makeewsrequestasync"></a>makeEwsRequestAsync の認証とアクセス許可について

このメソッドを `makeEwsRequestAsync` 使用すると、現在のユーザーの電子メール アカウント資格情報を使用して要求が認証されます。 このメソッドは `makeEwsRequestAsync` 、要求に認証資格情報を提供する必要がないように、資格情報を管理します。

> [!NOTE]
> サーバー管理者は [、New-WebServicesVirtualDirectory](/powershell/module/exchange/client-access-servers/New-WebServicesVirtualDirectory?view=exchange-ps&preserve-view=true) または [Set-WebServicesVirtualDirectory](/powershell/module/exchange/client-access-servers/Set-WebServicesVirtualDirectory?view=exchange-ps&preserve-view=true) コマンドレットを使用して、クライアント アクセス サーバー EWS ディレクトリで _OAuthAuthentication_ パラメーターを **true** に設定して、EWS 要求を行うメソッドを有効にする `makeEwsRequestAsync` 必要があります。

アドインで、メソッドを `ReadWriteMailbox` 使用 `makeEwsRequestAsync` するには、アドイン マニフェストでアクセス許可を指定する必要があります。 アクセス許可の`ReadWriteMailbox`使用については、「[Outlook アドイン](understanding-outlook-add-in-permissions.md)のアクセス許可について」の [ReadWriteMailbox アクセス許可](understanding-outlook-add-in-permissions.md#readwritemailbox-permission)に関するセクションを参照してください。

## <a name="see-also"></a>関連項目

- [Office アドインのプライバシーとセキュリティ](../concepts/privacy-and-security.md)
- [Office アドインにおける同一生成元ポリシーの制限への対処](../develop/addressing-same-origin-policy-limitations.md)
- 
  [Exchange 用 EWS リファレンス](/exchange/client-developer/web-service-reference/ews-reference-for-exchange)
- 
  [Exchange での Outlook 用メール アプリと EWS](/exchange/client-developer/exchange-web-services/mail-apps-for-outlook-and-ews-in-exchange)

ASP.NET Web APIを使用してアドインのバックエンド サービスを作成する方法については、次を参照してください。

- [ASP.NET Web API を使用して Office アドイン用 Web サービスを作成する](/archive/blogs/officeapps/create-a-web-service-for-an-app-for-office-using-the-asp-net-web-api)
- [ASP.NET Web API を使用した HTTP サービスの構築に関する基本](https://dotnet.microsoft.com/apps/aspnet/apis)
