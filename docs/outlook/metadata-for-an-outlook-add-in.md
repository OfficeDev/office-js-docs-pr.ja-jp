---
title: Outlook アドインでメタデータを取得および設定する
description: ローミング設定またはカスタム プロパティを使用して、Outlook アドインでカスタム データを管理します。
ms.date: 07/08/2022
ms.localizationpriority: medium
ms.openlocfilehash: b2cbb79288f7e62de8b4baae164ec9747cb83190
ms.sourcegitcommit: b6a3815a1ad17f3522ca35247a3fd5d7105e174e
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/22/2022
ms.locfileid: "66959050"
---
# <a name="get-and-set-add-in-metadata-for-an-outlook-add-in"></a>Outlook アドインのアドイン メタデータを取得および設定する

次のいずれかの方法を使用して、Outlook アドインでカスタム データを管理できます。

- ユーザーのメールボックスのカスタム データを管理するローミング設定。
- ユーザーのメールボックス内のアイテムのカスタム データを管理するカスタム プロパティ。

これらの方法の両方とも Outlook アドインでのみアクセス可能なカスタム データに対するアクセスを提供しますが、各方法は他方の方法と異なる方法でデータを保存します。つまり、ローミング設定によって保存されたデータはカスタム プロパティではアクセスできず、その逆もまた同様です。データは対象のメールボックスのサーバー に保存され、アドインでサポートされるすべてのフォーム ファクターのその後の Outlook セッションでアクセスできます。

## <a name="custom-data-per-mailbox-roaming-settings"></a>メールボックスごとのカスタム データ: ローミング設定

[RoamingSettings](/javascript/api/outlook/office.roamingsettings) オブジェクトを使用して、ユーザーの Exchange メールボックスに固有のデータを指定できます。このタイプのデータには、たとえばユーザーの個人データや基本設定があります。メール アドインは、その実行を許可されているデバイス (デスクトップ、タブレット、またはスマートフォン) でローミングするときにローミング設定にアクセスできます。

このデータへの変更は、現在の Outlook セッションの設定値のメモリ内コピーに格納されます。更新後にすべてのローミング設定値を明示的に保存して、ユーザーが次にアドインを同じデバイスで開いても、サポートされている他のデバイスで開いても、その設定値を使用できるようにしてください。

### <a name="roaming-settings-format"></a>ローミング設定の形式

**RoamingSettings** オブジェクトのデータは、シリアル化された JavaScript Object Notation (JSON) 文字列として格納されます。

`add-in_setting_name_0`、`add-in_setting_name_1`、`add-in_setting_name_2` という名前の 3 つのローミング設定が定義されていることを前提として、構造の例を次に示します。

```json
{
  "add-in_setting_name_0": "add-in_setting_value_0",
  "add-in_setting_name_1": "add-in_setting_value_1",
  "add-in_setting_name_2": "add-in_setting_value_2"
}
```

### <a name="loading-roaming-settings"></a>ローミング設定の読み込み

通常、メール アドインでは、[Office.initialize](/javascript/api/office#Office_initialize_reason_) イベント ハンドラーでローミング設定を読み込みます。 次の JavaScript コード例は、既存のローミング設定を読み込み、 **customerName** と **customerBalance** という 2 つの設定の値を取得する方法を示しています。

```js
let _mailbox;
let _settings;
let _customerName;
let _customerBalance;

// The initialize function is required for all add-ins.
Office.initialize = function () {
  // Initialize instance variables to access API objects.
  _mailbox = Office.context.mailbox;
  _settings = Office.context.roamingSettings;
  _customerName = _settings.get("customerName");
  _customerBalance = _settings.get("customerBalance");
}
```

### <a name="creating-or-assigning-a-roaming-setting"></a>ローミング設定の作成または割り当て

前の例の続きで、次の JavaScript 関数 `setAddInSetting` は、[RoamingSettings.set](/javascript/api/outlook/office.roamingsettings) メソッドを使用して `cookie` という名前の設定に今日の日付を設定し、[RoamingSettings.saveAsync](/javascript/api/outlook/office.roamingsettings#outlook-office-roamingsettings-saveasync-member(1)) メソッドを使用してすべてのローミング設定をサーバーに保存することによってデータを保存します。

このメソッドは `set` 、設定がまだ存在しない場合に設定を作成し、指定した値に設定を割り当てます。 このメソッドは `saveAsync` 、ローミング設定を非同期的に保存します。 このコード サンプルは、 `saveMyAddInSettingsCallback`コールバック関数を `saveAsync` 渡します。非同期呼び出しが終了すると、  `saveMyAddInSettingsCallback` _asyncResult_ という 1 つのパラメーターを使用して呼び出されます。 このパラメーターは [AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトであり、非同期呼び出しについての結果と詳細情報が格納されています。 オプションの _userContext_ パラメーターを使用すると、非同期呼び出しからコールバック関数に任意の状態情報を渡すことができます。

```js
// Set a roaming setting.
function setAddInSetting() {
  _settings.set("cookie", Date());
  // Save roaming settings for the mailbox
  // to the server so that they will be available
  // in the next session.
  _settings.saveAsync(saveMyAddInSettingsCallback);
}

// Callback function after saving custom roaming settings.
function saveMyAddInSettingsCallback(asyncResult) {
  if (asyncResult.status == Office.AsyncResultStatus.Failed) {
    // Handle the failure.
  }
}
```

### <a name="removing-a-roaming-setting"></a>ローミング設定の削除

さらに、前の例の続きで、次の JavaScript 関数  `removeAddInSetting` では、 [RoamingSettings.remove](/javascript/api/outlook/office.roamingsettings#outlook-office-roamingsettings-remove-member(1)) メソッドを使用して `cookie` 設定を削除し、すべてのローミング設定を Exchange Server に戻して保存する方法を示します。

```js
// Remove an add-in setting.
function removeAddInSetting()
{
  _settings.remove("cookie");
  // Save changes to the roaming settings for the mailbox
  // to the server so that they will be available
  // in the next session.
  _settings.saveAsync(saveMyAddInSettingsCallback);
}
```

## <a name="custom-data-per-item-in-a-mailbox-custom-properties"></a>メールボックス内のアイテムごとのカスタム データ: カスタム プロパティ

[CustomProperties](/javascript/api/outlook/office.customproperties) オブジェクトを使用して、ユーザーのメールボックス内のアイテムに固有のデータを指定することもできます。たとえば、メール アドインで特定のメッセージを分類し、カスタム プロパティ `messageCategory` を使用してそのカテゴリのメモを付けることができます。または、メール アドインでメッセージ内の会議の提案から予定を作成した場合に、カスタム プロパティを使用してそれぞれの予定を追跡できます。これにより、ユーザーが再度そのメッセージを開いた場合に、メール アドインによってもう一度予定を作成するように提案されることはありません。

ローミング設定と同様に、カスタム プロパティに対する変更は現在の Outlook セッションのプロパティのメモリ内コピーに格納されます。これらのカスタム プロパティが次のセッションで使用できるようにするには、[CustomProperties.saveAsync](/javascript/api/outlook/office.customproperties#outlook-office-customproperties-saveasync-member(1))を使用します。

これらのアドイン固有のアイテム固有のカスタム プロパティには、オブジェクトを使用 `CustomProperties` してのみアクセスできます。 これらのプロパティは、Outlook オブジェクト モデルのカスタム MAPI ベースの [UserProperties](/office/vba/api/Outlook.UserProperties) と、Exchange Web Services (EWS) の拡張プロパティとは異なります。 Outlook オブジェクト モデル、EWS、または REST を使用して直接アクセス `CustomProperties` することはできません。 EWS または REST を使用してアクセス `CustomProperties` する方法については、「 [EWS または REST を使用してカスタム プロパティを取得する](#get-custom-properties-using-ews-or-rest)」セクションを参照してください。

### <a name="using-custom-properties"></a>カスタム プロパティの使用

カスタム プロパティを使用するには、まず [loadCustomPropertiesAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods) メソッドを呼び出して読み込む必要があります方法です。 プロパティ バッグを作成したら、[set](/javascript/api/outlook/office.customproperties#outlook-office-customproperties-set-member(1)) と [get](/javascript/api/outlook/office.customproperties) メソッドを使用してカスタム プロパティを追加し、取得できます。 プロパティ バッグで行った変更を保存するには、[saveAsync](/javascript/api/outlook/office.customproperties#outlook-office-customproperties-saveasync-member(1)) メソッドを使用する必要があります。

 > [!NOTE]
 > Outlook on Mac では、カスタム プロパティをキャッシュに入れないため、ユーザーのネットワークがダウンした場合、Outlook on Mac のメール アドインでカスタム プロパティにアクセスできなくなります。

### <a name="custom-properties-example"></a>カスタム プロパティの例

次の例は、カスタム プロパティを使用する Outlook アドインの関数とメソッドの簡略化されたセットを示しています。 この例を出発点として、カスタム プロパティを使用するアドインを作成できます。

この例には、次の関数とメソッドが含まれます。

- [Office.initialize](/javascript/api/office#Office_initialize_reason_) -- アドインを初期化し、Exchange Server からカスタム プロパティ バッグを読み込みます。

- **customPropsCallback** -- サーバーから返されるカスタム プロパティ バッグを取得し、後で使用できるように保存します。

- **updateProperty** -- 特定のプロパティを設定または更新し、変更をサーバーに保存します。

- **removeProperty** -- プロパティ バッグから特定のプロパティを削除し、その削除をサーバーに保存します。

```js
let _mailbox;
let _customProps;

// The initialize function is required for all add-ins.
Office.initialize = function () {
  _mailbox = Office.context.mailbox;
  _mailbox.item.loadCustomPropertiesAsync(customPropsCallback);
}

// Callback function from loading custom properties.
function customPropsCallback(asyncResult) {
  if (asyncResult.status == Office.AsyncResultStatus.Failed) {
    // Handle the failure.
  }
  else {
    // Successfully loaded custom properties,
    // can get them from the asyncResult argument.
    _customProps = asyncResult.value;
  }
}

// Get individual custom property.
function getProperty() {
  const myProp = _customProps.get("myProp");
}

// Set individual custom property.
function updateProperty(name, value) {
  _customProps.set(name, value);
  // Save all custom properties to server.
  _customProps.saveAsync(saveCallback);
}

// Remove a custom property.
function removeProperty(name) {
  _customProps.remove(name);
  // Save all custom properties to server.
  _customProps.saveAsync(saveCallback);
}

// Callback function from saving custom properties.
function saveCallback() {
  if (asyncResult.status == Office.AsyncResultStatus.Failed) {
    // Handle the failure.
  }
}
```

### <a name="get-custom-properties-using-ews-or-rest"></a>EWS または REST を使用してカスタム プロパティを取得する

EWS または REST を使用して **CustomProperties** を取得する場合は、最初にMAPI ベースの拡張プロパティの名前を決めるようにします。 その後、MAPI ベースの拡張プロパティを取得するのと同じ方法でそのプロパティを取得できます。

#### <a name="how-custom-properties-are-stored-on-an-item"></a>アイテムでのカスタム プロパティの格納方法

アドインによって設定されたカスタム プロパティは、標準の MAPI ベースのプロパティとは異なります。 アドイン API は、すべてのアドインを JSON ペイロードとしてシリアル化し、名前 (アドインの `CustomProperties` ID) とプロパティ セット GUID `{00020329-0000-0000-C000-000000000046}`を持`cecp-<app-guid>`つ単一の MAPI ベースの拡張プロパティに保存`<app-guid>`します。 (このオブジェクトに関する詳細については、「[MS OXCEXT 2.2.5 メール アプリのカスタム プロパティ](/openspecs/exchange_server_protocols/ms-oxcext/4cf1da5e-c68e-433e-a97e-c45625483481)」を参照してください。) その後、EWS または REST を使用してこの MAPI ベースのプロパティを取得できます。

#### <a name="get-custom-properties-using-ews"></a>EWS を使用してカスタム プロパティを取得する

メール アドインは、EWS [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) 操作を`CustomProperties`使用して MAPI ベースの拡張プロパティを取得できます。 コールバック トークンを使用してサーバー側にアクセスするか、[mailbox.makeEwsRequestAsync メソッドを](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods)使用してクライアント側でアクセス`GetItem`します。 要求で、前の `GetItem` セクション「カスタム プロパティを `CustomProperties` アイテムに格納する方法」で説明した詳細を使用して、そのプロパティ セットで MAPI ベース [のプロパティを指定します](#how-custom-properties-are-stored-on-an-item)。

次の例では、アイテムとそれのカスタム プロパティを取得する方法を示します。

> [!IMPORTANT]
> 次の例では、`<app-guid>` を自分のアドインの ID と置き換えてください。

```typescript
let request_str =
    '<?xml version="1.0" encoding="utf-8"?>' +
    '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"' +
                   'xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"' +
                   'xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"' +
                   'xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">' +
        '<soap:Header xmlns:wsse="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-secext-1.0.xsd"' +
                     'xmlns:wsa="http://www.w3.org/2005/08/addressing">' +
            '<t:RequestServerVersion Version="Exchange2010_SP1"/>' +
        '</soap:Header>' +
        '<soap:Body>' +
            '<m:GetItem>' +
                '<m:ItemShape>' +
                    '<t:BaseShape>AllProperties</t:BaseShape>' +
                    '<t:IncludeMimeContent>true</t:IncludeMimeContent>' +
                    '<t:AdditionalProperties>' +
                        '<t:ExtendedFieldURI ' +
                          'DistinguishedPropertySetId="PublicStrings" ' +
                          'PropertyName="cecp-<app-guid>"' +
                          'PropertyType="String" ' +
                        '/>' +
                    '</t:AdditionalProperties>' +
                '</m:ItemShape>' +
                '<m:ItemIds>' +
                    '<t:ItemId Id="' +
                      Office.context.mailbox.item.itemId +
                    '"/>' +
                '</m:ItemIds>' +
            '</m:GetItem>' +
        '</soap:Body>' +
    '</soap:Envelope>';

Office.context.mailbox.makeEwsRequestAsync(
    request_str,
    function(asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
            console.log(asyncResult.value);
        }
        else {
            console.log(JSON.stringify(asyncResult));
        }
    }
);
```

要求文字列で他のカスタム プロパティを [ExtendedFieldURI](/exchange/client-developer/web-service-reference/extendedfielduri) 要素として指定すると、それらのカスタム プロパティを取得することができます。

#### <a name="get-custom-properties-using-rest"></a>REST を使用してカスタム プロパティを取得する

アドインでメッセージやイベントに対して REST クエリを作成し、すでにカスタム プロパティを持つメッセージやイベントを取得することができます。 「[アイテムでのカスタム プロパティの格納方法](#how-custom-properties-are-stored-on-an-item)」セクションで説明されている詳細を参考にして、クエリに MAPI ベースの拡張プロパティ **CustomProperties** とそのプロパティ セットを含めます。

次の例では、アドインで設定されたいずれかのカスタム プロパティを含むすべてのイベントを取得し、追加のフィルター処理のロジックを適用できるようにプロパティの値を応答に含める方法を示します。

> [!IMPORTANT]
> 次の例では、`<app-guid>` を自分のアドインの ID と置き換えてください。

```rest
GET https://outlook.office.com/api/v2.0/Me/Events?$filter=SingleValueExtendedProperties/Any
  (ep: ep/PropertyId eq 'String {00020329-0000-0000-C000-000000000046}
  Name cecp-<app-guid>' and ep/Value ne null)
  &$expand=SingleValueExtendedProperties($filter=PropertyId eq 'String
  {00020329-0000-0000-C000-000000000046} Name cecp-<app-guid>')
```

REST を使用して単一値の MAPI ベースの拡張プロパティを取得するその他の例は、「[singleValueExtendedProperty を取得る](/graph/api/singlevaluelegacyextendedproperty-get?view=graph-rest-1.0&preserve-view=true)」 を参照してください。

次の例では、アイテムとそれのカスタム プロパティを取得する方法を示します。 `done` メソッドのコールバック関数では、要求されたカスタム プロパティの一覧は `item.SingleValueExtendedProperties` に含まれます。

> [!IMPORTANT]
> 次の例では、`<app-guid>` を自分のアドインの ID と置き換えてください。

```typescript
Office.context.mailbox.getCallbackTokenAsync(
    {
        isRest: true
    },
    function (asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded
            && asyncResult.value !== "") {
            let item_rest_id = Office.context.mailbox.convertToRestId(
                Office.context.mailbox.item.itemId,
                Office.MailboxEnums.RestVersion.v2_0);
            let rest_url = Office.context.mailbox.restUrl +
                           "/v2.0/me/messages('" +
                           item_rest_id +
                           "')";
            rest_url += "?$expand=SingleValueExtendedProperties($filter=PropertyId eq 'String {00020329-0000-0000-C000-000000000046} Name cecp-<app-guid>')";

            let auth_token = asyncResult.value;
            $.ajax(
                {
                    url: rest_url,
                    dataType: 'json',
                    headers:
                        {
                            "Authorization":"Bearer " + auth_token
                        }
                }
                ).done(
                    function (item) {
                        console.log(JSON.stringify(item));
                    }
                ).fail(
                    function (error) {
                        console.log(JSON.stringify(error));
                    }
                );
        } else {
            console.log(JSON.stringify(asyncResult));
        }
    }
);
```

## <a name="see-also"></a>関連項目

- [MAPI のプロパティの概要](/office/client-developer/outlook/mapi/mapi-property-overview)
- [Outlook のプロパティの概要](/office/vba/outlook/How-to/Navigation/properties-overview)  
- [Outlook アドインからの Outlook REST API の呼び出し](use-rest-api.md)
- [Outlook アドインから Web サービスを呼び出す](web-services.md)
- 
  [Exchange における EWS のプロパティと拡張プロパティ](/exchange/client-developer/exchange-web-services/properties-and-extended-properties-in-ews-in-exchange)
- 
  [Exchange の EWS でのプロパティ セットと応答の図形](/exchange/client-developer/exchange-web-services/property-sets-and-response-shapes-in-ews-in-exchange)
