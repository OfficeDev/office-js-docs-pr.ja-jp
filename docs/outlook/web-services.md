---
title: Outlook アドインから Exchange Web サービス (EWS) を使用する
description: Outlook アドインが Exchange Web サービスに情報を要求する方法の例を示します。
ms.date: 10/31/2019
localization_priority: Normal
ms.openlocfilehash: 4c0c97a9a796dc1f257b1a0b0ec880b3ca3d8e74
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/20/2020
ms.locfileid: "42166442"
---
# <a name="call-web-services-from-an-outlook-add-in"></a>Outlook アドインから Web サービスを呼び出す

アドインは、Exchange Server 2013 が実行されているコンピューター、アドインの UI のソースの場所を提供するサーバー上で利用できる Web サービス、またはインターネット上で利用できる Web サービスから Exchange Web サービス (EWS) を使用できます。この記事では、Outlook アドインからどのように EWS の情報を要求できるかを示す例を説明します。

Web サービスを呼び出す方法は、Web サービスは配置された場所によって異なります。表 1 は、場所により異なる Web サービスを呼び出すさまざまな方法を示しています。


**表 1.Outlook アドインから web サービスを呼び出す方法**

<br/>

|**Web サービスの場所**|**Web サービスを呼び出す方法**|
|:-----|:-----|
|クライアント メールボックスをホストする Exchange サーバー|[mailbox.makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) メソッドを使用して、アドインがサポートしている EWS 操作を呼び出します。また、メールボックスをホストしている Exchange サーバーも EWS を公開します。|
|アドインの UI のソースの場所を提供する Web サーバー|標準の JavaScript の手法を使用して Web サービスを呼び出します。UI フレーム内の JavaScript コードは、UI を提供する Web サーバーのコンテキストで実行されます。そのため、クロスサイト スクリプト エラーを発生させることなく、そのサーバーで Web サービスを呼び出すことができます。|
|上記以外の場所|UI のソースの場所を提供する Web サーバー上で、Web サービスのプロキシを作成します。プロキシを作成しないと、クロスサイト スクリプト エラーによってアドインを実行できなくなります。プロキシを作成する方法の 1 つは JSON/P を使用することです。詳細については、「 [Office アドインのプライバシーとセキュリティ](../develop/privacy-and-security.md)」を参照してください。|

## <a name="using-the-makeewsrequestasync-method-to-access-ews-operations"></a>makeEwsRequestAsync メソッドを使用して EWS 操作にアクセスする

[mailbox.makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) メソッドを使用して、ユーザーのメールボックスをホストしている Exchange サーバーに EWS 要求を行うことができます。

EWS は、Exchange サーバーでの種々の操作をサポートしています。たとえば、アイテム レベルの操作としては、アイテムのコピー、検索、更新、または送信などがあり、フォルダー レベルの操作としては、フォルダーの作成、取得、または更新などがあります。EWS 操作を実行するには、その操作の XML SOAP 要求を作成します。操作が終了すると、その操作に関係するデータが含まれた XML SOAP 応答を受信します。EWS SOAP の要求と応答は、Messages.xsd ファイルで定義されているスキーマに従います。Messages.xsd ファイルは、他の EWS スキーマ ファイルと同様、EWS をホストしている IIS 仮想ディレクトリに配置されています。

**makeEwsRequestAsync** メソッドを使用して EWS 操作を開始するには、以下の項目を指定します。

- その EWS 操作の SOAP 要求に対する XML ( _data_ パラメーターへの引数)

- コールバック メソッド ( _callback_ 引数)

- そのコールバック メソッドに対するオプションの入力データ ( _userContext_ 引数)

EWS SOAP 要求が完了すると、Outlook が 1 つの引数 ([AsyncResult](/javascript/api/office/office.asyncresult) オブジェクト) を使用してコールバック メソッドを呼び出します。コールバック メソッドは、 **AsyncResult** オブジェクトの 2 つのプロパティにアクセスできます。1 つは **value** プロパティ (EWS 操作の XML SOAP 応答が含まれる)、もう 1 つはオプションの **asyncContext** プロパティ ( **userContext** パラメーターとして渡したデータが含まれる) です。通常、コールバック メソッドは、SOAP 応答内の XML を解析して、関係する情報を取得し、その情報を処理します。


## <a name="tips-for-parsing-ews-responses"></a>EWS 応答を解析するためのヒント

SOAP 応答を EWS 操作から解析する場合、ブラウザーに依存する以下の問題に注意してください。


- DOM メソッド **getElementsByTagName** を使用する場合、Internet Explorer のサポートを組み込むため、タグ名に接頭辞を指定します。

  **getElementsByTagName** の動作は、ブラウザーの種類によって異なります。たとえば、EWS の応答には、次の XML (表示目的で書式設定され省略されています) を含めることができます。

   ```XML
        <t:ExtendedProperty><t:ExtendedFieldURI PropertySetId="00000000-0000-0000-0000-000000000000" 
        PropertyName="MyProperty" 
        PropertyType="String"/>
        <t:Value>{
        ...
        }</t:Value></t:ExtendedProperty>
   ```

   以下に挙げるコードは、Chrome などのブラウザーで **ExtendedProperty** タグで囲まれた XML を取得するために使用できます。

   ```js
        var mailbox = Office.context.mailbox;
        mailbox.makeEwsRequestAsync(mailbox.item.itemId, function(result) {
            var response = $.parseXML(result.value);
            var extendedProps = response.getElementsByTagName("ExtendedProperty")
            });
   ```

   Internet Explorer では、以下に示すように、タグ名に接頭辞 `t:` を含める必要があります。

   ```js
        var mailbox = Office.context.mailbox;
        mailbox.makeEwsRequestAsync(mailbox.item.itemId, function(result) {
            var response = $.parseXML(result.value);
            var extendedProps = response.getElementsByTagName("t:ExtendedProperty")
            });
   ```

- 以下のように DOM プロパティ **textContent** を使用し、EWS 応答のタグの内容を取得します。
    
   ```js
      content = $.parseJSON(value.textContent);
   ```

   Internet Explorer では、**innerHTML** などのその他のプロパティは EWS 応答のいくつかのタグに対して機能しない場合があります。
    

## <a name="example"></a>例

次の例では、**makeEwsRequestAsync** を呼び出して、[GetItem](/exchange/client-developer/web-service-reference/getitem-operation) 操作を使用してアイテムの件名を取得します。この例には、次の 3 つの関数が含まれています。

-  `getSubjectRequest` &ndash; 入力としてアイテム ID を受け取り、指定のアイテムに対する **GetItem** を呼び出すための SOAP 要求の XML を返します。
    
-  `sendRequest` &ndash; `getSubjectRequest` を呼び出して指定のアイテムに対する SOAP 要求を取得し、SOAP 要求とコールバック メソッド (`callback`) を **makeEwsRequestAsync** に渡して、指定のアイテムの件名を取得します。
    
-  `callback` &ndash; 指定のアイテムの件名とその他の情報が含まれている SOAP 応答を処理します。
    

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


## <a name="ews-operations-that-add-ins-support"></a>アドインでサポートしている EWS 操作

Outlook アドインでは、**makeEwsRequestAsync** メソッドを使用して、EWS で利用できる操作のサブセットにアクセスできます。EWS 操作や、**makeEwsRequestAsync** メソッドを使用して操作にアクセスする方法が不明な場合は、はじめに _data_ 引数をカスタマイズする SOAP 要求の例を確認してください。 

以下に、**makeEwsRequestAsync** メソッドを使用する方法を説明します。

1. XML 内のアイテム ID および関係する EWS 操作属性を適切な値に置き換えます。
    
2. _makeEwsRequestAsync_ の **data** パラメーターの引数として SOAP 要求を指定します。
    
3. コールバック メソッドを指定し、**makeEwsRequestAsync** を呼び出します。
    
4. コールバック メソッド内で、SOAP 応答内の操作の結果を検証します。
    
5. 必要に応じて EWS 操作の結果を使用します。
    
次の表は、アドインがサポートしている EWS 操作を示しています。SOAP の要求と応答の例を表示するには、各操作のリンクを選択します。EWS 操作の詳細については、「 [Exchange での EWS の操作](/exchange/client-developer/web-service-reference/ews-operations-in-exchange)」を参照してください。

**表 2サポートされている EWS 操作**

<br/>

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

**makeEwsRequestAsync** メソッドを使用すると、現在のユーザーの電子メール アカウントの資格情報を使用して要求が認証されます。 資格情報は **makeEwsRequestAsync** メソッドが管理するため、要求と共に認証資格情報を提供する必要はありません。

> [!NOTE]
> サーバー管理者は、**makeEwsRequestAsync** メソッドで EWS 要求を送信できるようにするため、[New-WebServicesVirtualDirectory](/powershell/module/exchange/client-access-servers/New-WebServicesVirtualDirectory?view=exchange-ps) コマンドレットまたは [Set-WebServicesVirtualDirectory](/powershell/module/exchange/client-access-servers/Set-WebServicesVirtualDirectory?view=exchange-ps) コマンドレットを使用して、クライアント アクセス サーバーの EWS ディレクトリで _OAuthAuthentication_ パラメーターを **true** にする必要があります。

**makeEwsRequestAsync** メソッドを使用するには、アドインのマニフェストで **ReadWriteMailbox** アクセス許可を指定する必要があります。 **ReadWriteMailbox** アクセス許可の使用については、「[Outlook アドインのアクセス許可を理解する](understanding-outlook-add-in-permissions.md)」の「[ReadWriteMailbox アクセス許可](understanding-outlook-add-in-permissions.md#readwritemailbox-permission)」セクションを参照してください。

> [!NOTE]
> サーバー管理者は、**makeEwsRequestAsync** メソッドで EWS 要求を送信できるようにするため、[New-WebServicesVirtualDirectory](/powershell/module/exchange/client-access-servers/New-WebServicesVirtualDirectory?view=exchange-ps) コマンドレットまたは [Set-WebServicesVirtualDirectory](/powershell/module/exchange/client-access-servers/Set-WebServicesVirtualDirectory?view=exchange-ps) コマンドレットを使用して、クライアント アクセス サーバーの EWS ディレクトリで _OAuthAuthentication_ パラメーターを **true** にする必要があります。



## <a name="see-also"></a>関連項目

- [Office アドインのプライバシーとセキュリティ](../develop/privacy-and-security.md)   
- [Office アドインにおける同一生成元ポリシーの制限への対処](../develop/addressing-same-origin-policy-limitations.md)
- 
  [Exchange 用 EWS リファレンス](/exchange/client-developer/web-service-reference/ews-reference-for-exchange)   
- 
  [Exchange での Outlook 用メール アプリと EWS](/exchange/client-developer/exchange-web-services/mail-apps-for-outlook-and-ews-in-exchange)
   
ASP.NET Web API を使用してアドイン用のバックエンド サービスを作成する場合は、以下の資料を参照してください。

- [ASP.NET Web API を使用して Office アドイン用 Web サービスを作成する](https://blogs.msdn.microsoft.com/officeapps/2013/06/10/create-a-web-service-for-an-app-for-office-using-the-asp-net-web-api/)    
- [ASP.NET Web API を使用した HTTP サービスの構築に関する基本](https://www.asp.net/web-api)
    
