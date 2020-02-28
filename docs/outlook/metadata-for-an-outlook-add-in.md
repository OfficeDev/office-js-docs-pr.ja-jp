---
title: Outlook アドインでメタデータを取得および設定する
description: ローミング設定またはカスタム プロパティを使用して、Outlook アドインでカスタム データを管理します。
ms.date: 10/31/2019
localization_priority: Normal
ms.openlocfilehash: 3bf19f56b11b524ea2ee722e2997465bbd36d55c
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/27/2020
ms.locfileid: "42324934"
---
# <a name="get-and-set-add-in-metadata-for-an-outlook-add-in"></a><span data-ttu-id="c04ce-103">Outlook アドインのアドイン メタデータを取得および設定する</span><span class="sxs-lookup"><span data-stu-id="c04ce-103">Get and set add-in metadata for an Outlook add-in</span></span>

<span data-ttu-id="c04ce-104">次のいずれかの方法を使用して、Outlook アドインでカスタム データを管理できます。</span><span class="sxs-lookup"><span data-stu-id="c04ce-104">You can manage custom data in your Outlook add-in by using either of the following:</span></span>

- <span data-ttu-id="c04ce-105">ユーザーのメールボックスのカスタム データを管理するローミング設定。</span><span class="sxs-lookup"><span data-stu-id="c04ce-105">Roaming settings, which manage custom data for a user's mailbox.</span></span>
- <span data-ttu-id="c04ce-106">ユーザーのメールボックス内のアイテムのカスタム データを管理するカスタム プロパティ。</span><span class="sxs-lookup"><span data-stu-id="c04ce-106">Custom properties, which manage custom data for an item in a user's mailbox.</span></span>

<span data-ttu-id="c04ce-p101">これらの方法の両方とも Outlook アドインでのみアクセス可能なカスタム データに対するアクセスを提供しますが、各方法は他方の方法と異なる方法でデータを保存します。つまり、ローミング設定によって保存されたデータはカスタム プロパティではアクセスできず、その逆もまた同様です。データは対象のメールボックスのサーバー に保存され、アドインでサポートされるすべてのフォーム ファクターのその後の Outlook セッションでアクセスできます。</span><span class="sxs-lookup"><span data-stu-id="c04ce-p101">Both of these give access to custom data that is only accessible by your Outlook add-in, but each method stores the data separately from the other. That is, the data stored through roaming settings is not accessible by custom properties, and vice versa. The data is stored on the server for that mailbox, and is accessible in subsequent Outlook sessions on all the form factors that the add-in supports.</span></span>

## <a name="custom-data-per-mailbox-roaming-settings"></a><span data-ttu-id="c04ce-110">メールボックスごとのカスタム データ: ローミング設定</span><span class="sxs-lookup"><span data-stu-id="c04ce-110">Custom data per mailbox: roaming settings</span></span>

<span data-ttu-id="c04ce-p102">[RoamingSettings](/javascript/api/outlook/office.RoamingSettings) オブジェクトを使用して、ユーザーの Exchange メールボックスに固有のデータを指定できます。このタイプのデータには、たとえばユーザーの個人データや基本設定があります。メール アドインは、その実行を許可されているデバイス (デスクトップ、タブレット、またはスマートフォン) でローミングするときにローミング設定にアクセスできます。</span><span class="sxs-lookup"><span data-stu-id="c04ce-p102">You can specify data specific to a user's Exchange mailbox using the [RoamingSettings](/javascript/api/outlook/office.RoamingSettings) object. Examples of such data include the user's personal data and preferences. Your mail add-in can access roaming settings when it roams on any device it's designed to run on (desktop, tablet, or smartphone).</span></span>

<span data-ttu-id="c04ce-p103">このデータへの変更は、現在の Outlook セッションの設定値のメモリ内コピーに格納されます。更新後にすべてのローミング設定値を明示的に保存して、ユーザーが次にアドインを同じデバイスで開いても、サポートされている他のデバイスで開いても、その設定値を使用できるようにしてください。</span><span class="sxs-lookup"><span data-stu-id="c04ce-p103">Changes to this data are stored on an in-memory copy of those settings for the current Outlook session. You should explicitly save all the roaming settings after updating them so that they will be available the next time the user opens your add-in, on the same or any other supported device.</span></span>


### <a name="roaming-settings-format"></a><span data-ttu-id="c04ce-116">ローミング設定の形式</span><span class="sxs-lookup"><span data-stu-id="c04ce-116">Roaming settings format</span></span>

<span data-ttu-id="c04ce-117">**RoamingSettings** オブジェクトのデータは、シリアル化された JavaScript Object Notation (JSON) 文字列として格納されます。</span><span class="sxs-lookup"><span data-stu-id="c04ce-117">The data in a **RoamingSettings** object is stored as a serialized JavaScript Object Notation (JSON) string.</span></span> 

<span data-ttu-id="c04ce-118">`add-in_setting_name_0`、`add-in_setting_name_1`、`add-in_setting_name_2` という名前の 3 つのローミング設定が定義されていることを前提として、構造の例を次に示します。</span><span class="sxs-lookup"><span data-stu-id="c04ce-118">The following is an example of the structure, assuming there are three defined roaming settings named `add-in_setting_name_0`,  `add-in_setting_name_1`, and  `add-in_setting_name_2`.</span></span>


```json
{
  "add-in_setting_name_0": "add-in_setting_value_0",
  "add-in_setting_name_1": "add-in_setting_value_1",
  "add-in_setting_name_2": "add-in_setting_value_2"
}
```


### <a name="loading-roaming-settings"></a><span data-ttu-id="c04ce-119">ローミング設定の読み込み</span><span class="sxs-lookup"><span data-stu-id="c04ce-119">Loading roaming settings</span></span>

<span data-ttu-id="c04ce-120">通常、メール アドインでは、[Office.initialize](/javascript/api/office#office-initialize-reason-) イベント ハンドラーでローミング設定を読み込みます。</span><span class="sxs-lookup"><span data-stu-id="c04ce-120">A mail add-in typically loads roaming settings in the [Office.initialize](/javascript/api/office#office-initialize-reason-) event handler.</span></span> <span data-ttu-id="c04ce-121">次の JavaScript コードは、既存のローミング設定を読み込み、2 つの設定 **customerName** と **customerBalance** の値を取得する例を示しています。</span><span class="sxs-lookup"><span data-stu-id="c04ce-121">The following JavaScript code example shows how to load existing roaming settings and get the values of 2 settings, **customerName** and **customerBalance**:</span></span>


```js
var _mailbox;
var _settings;
var _customerName;
var _customerBalance;

// The initialize function is required for all add-ins.
Office.initialize = function () {
  // Initialize instance variables to access API objects.
  _mailbox = Office.context.mailbox;
  _settings = Office.context.roamingSettings;
  _customerName = _settings.get("customerName");
  _customerBalance = _settings.get("customerBalance");
}

```


### <a name="creating-or-assigning-a-roaming-setting"></a><span data-ttu-id="c04ce-122">ローミング設定の作成または割り当て</span><span class="sxs-lookup"><span data-stu-id="c04ce-122">Creating or assigning a roaming setting</span></span>

<span data-ttu-id="c04ce-123">前の例の続きで、次の JavaScript 関数 `setAddInSetting` は、[RoamingSettings.set](/javascript/api/outlook/office.RoamingSettings) メソッドを使用して `cookie` という名前の設定に今日の日付を設定し、[RoamingSettings.saveAsync](/javascript/api/outlook/office.RoamingSettings#saveasync-callback-) メソッドを使用してすべてのローミング設定をサーバーに保存することによってデータを保存します。</span><span class="sxs-lookup"><span data-stu-id="c04ce-123">Continuing with the preceding example, the following JavaScript function,  `setAddInSetting`, shows how to use the [RoamingSettings.set](/javascript/api/outlook/office.RoamingSettings) method to set a setting named `cookie` with today's date, and persist the data by using the [RoamingSettings.saveAsync](/javascript/api/outlook/office.RoamingSettings#saveasync-callback-) method to save all the roaming settings back to the server.</span></span>

<span data-ttu-id="c04ce-124">この`set`設定が存在しない場合、メソッドは設定を作成し、指定された値に設定を割り当てます。</span><span class="sxs-lookup"><span data-stu-id="c04ce-124">The `set` method creates the setting if the setting does not already exist, and assigns the setting to the specified value.</span></span> <span data-ttu-id="c04ce-125">メソッド`saveAsync`は、ローミング設定を非同期的に保存します。</span><span class="sxs-lookup"><span data-stu-id="c04ce-125">The `saveAsync` method saves roaming settings asynchronously.</span></span> <span data-ttu-id="c04ce-126">このコードサンプル`saveMyAddInSettingsCallback` `saveMyAddInSettingsCallback`は、コールバックメソッドを渡し`saveAsync`ます。非同期呼び出しが完了すると、は1つのパラメーター _asyncResult_を使用して呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="c04ce-126">This code sample passes a callback method, `saveMyAddInSettingsCallback`, to `saveAsync` When the asynchronous call finishes,  `saveMyAddInSettingsCallback` is called by using one parameter, _asyncResult_.</span></span> <span data-ttu-id="c04ce-127">このパラメーターは [AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトであり、非同期呼び出しについての結果と詳細情報が格納されています。</span><span class="sxs-lookup"><span data-stu-id="c04ce-127">This parameter is an [AsyncResult](/javascript/api/office/office.asyncresult) object that contains the result of and any details about the asynchronous call.</span></span> <span data-ttu-id="c04ce-128">オプションの _userContext_ パラメーターを使用すると、非同期呼び出しからコールバック関数に任意の状態情報を渡すことができます。</span><span class="sxs-lookup"><span data-stu-id="c04ce-128">You can use the optional _userContext_ parameter to pass any state information from the asynchronous call to the callback function.</span></span>

```js
// Set a roaming setting.
function setAddInSetting() {
  _settings.set("cookie", Date());
  // Save roaming settings for the mailbox
  // to the server so that they will be available
  // in the next session.
  _settings.saveAsync(saveMyAddInSettingsCallback);
}

// Callback method after saving custom roaming settings.
function saveMyAddInSettingsCallback(asyncResult) {
  if (asyncResult.status == Office.AsyncResultStatus.Failed) {
    // Handle the failure.
  }
}
```


### <a name="removing-a-roaming-setting"></a><span data-ttu-id="c04ce-129">ローミング設定の削除</span><span class="sxs-lookup"><span data-stu-id="c04ce-129">Removing a roaming setting</span></span>

<span data-ttu-id="c04ce-130">さらに、前の例の続きで、次の JavaScript 関数  `removeAddInSetting` では、 [RoamingSettings.remove](/javascript/api/outlook/office.RoamingSettings#remove-name-) メソッドを使用して `cookie` 設定を削除し、すべてのローミング設定を Exchange Server に戻して保存する方法を示します。</span><span class="sxs-lookup"><span data-stu-id="c04ce-130">Also extending the preceding examples, the following JavaScript function,  `removeAddInSetting`, shows how to use the [RoamingSettings.remove](/javascript/api/outlook/office.RoamingSettings#remove-name-) method to remove the `cookie` setting and save all the roaming settings back to the Exchange Server.</span></span>


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


## <a name="custom-data-per-item-in-a-mailbox-custom-properties"></a><span data-ttu-id="c04ce-131">メールボックス内のアイテムごとのカスタム データ: カスタム プロパティ</span><span class="sxs-lookup"><span data-stu-id="c04ce-131">Custom data per item in a mailbox: custom properties</span></span>

<span data-ttu-id="c04ce-p106">[CustomProperties](/javascript/api/outlook/office.CustomProperties) オブジェクトを使用して、ユーザーのメールボックス内のアイテムに固有のデータを指定することもできます。たとえば、メール アドインで特定のメッセージを分類し、カスタム プロパティ `messageCategory` を使用してそのカテゴリのメモを付けることができます。または、メール アドインでメッセージ内の会議の提案から予定を作成した場合に、カスタム プロパティを使用してそれぞれの予定を追跡できます。これにより、ユーザーが再度そのメッセージを開いた場合に、メール アドインによってもう一度予定を作成するように提案されることはありません。</span><span class="sxs-lookup"><span data-stu-id="c04ce-p106">You can specify data specific to an item in the user's mailbox using the [CustomProperties](/javascript/api/outlook/office.CustomProperties) object. For example, your mail add-in could categorize certain messages and note the category using a custom property `messageCategory`. Or, if your mail add-in creates appointments from meeting suggestions in a message, you can use a custom property to track each of these appointments. This ensures that if the user opens the message again, your mail add-in doesn't offer to create the appointment a second time.</span></span>

<span data-ttu-id="c04ce-p107">ローミング設定と同様に、カスタム プロパティに対する変更は現在の Outlook セッションのプロパティのメモリ内コピーに格納されます。これらのカスタム プロパティが次のセッションで使用できるようにするには、[CustomProperties.saveAsync](/javascript/api/outlook/office.CustomProperties#saveasync-callback--asynccontext-)を使用します。</span><span class="sxs-lookup"><span data-stu-id="c04ce-p107">Similar to roaming settings, changes to custom properties are stored on in-memory copies of the properties for the current Outlook session. To make sure these custom properties will be available in the next session, use [CustomProperties.saveAsync](/javascript/api/outlook/office.CustomProperties#saveasync-callback--asynccontext-).</span></span>

<span data-ttu-id="c04ce-138">これらのアドイン固有のアイテム固有のカスタムプロパティにアクセスするには、その`CustomProperties`オブジェクトを使用する必要があります。</span><span class="sxs-lookup"><span data-stu-id="c04ce-138">These add-in-specific, item-specific custom properties can only be accessed by using the `CustomProperties` object.</span></span> <span data-ttu-id="c04ce-139">これらのプロパティは、Outlook オブジェクトモデルのカスタム、MAPI ベースの[UserProperties](/office/vba/api/Outlook.UserProperties) 、および Exchange Web サービス (EWS) の拡張プロパティとは異なります。</span><span class="sxs-lookup"><span data-stu-id="c04ce-139">These properties are different from the custom, MAPI-based [UserProperties](/office/vba/api/Outlook.UserProperties) in the Outlook object model, and extended properties in Exchange Web Services (EWS).</span></span> <span data-ttu-id="c04ce-140">Outlook オブジェクトモデル、 `CustomProperties` EWS、または REST を使用して直接アクセスすることはできません。</span><span class="sxs-lookup"><span data-stu-id="c04ce-140">You cannot directly access `CustomProperties` by using the Outlook object model, EWS, or REST.</span></span> <span data-ttu-id="c04ce-141">EWS または REST を`CustomProperties`使用してアクセスする方法については、「 [EWS または rest を使用してカスタムプロパティを取得](#get-custom-properties-using-ews-or-rest)する」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="c04ce-141">To learn how to access `CustomProperties` using EWS or REST, see the section [Get custom properties using EWS or REST](#get-custom-properties-using-ews-or-rest).</span></span>

### <a name="using-custom-properties"></a><span data-ttu-id="c04ce-142">カスタム プロパティの使用</span><span class="sxs-lookup"><span data-stu-id="c04ce-142">Using custom properties</span></span>

<span data-ttu-id="c04ce-143">カスタム プロパティを使用するには、まず [loadCustomPropertiesAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) メソッドを呼び出して読み込む必要があります方法です。</span><span class="sxs-lookup"><span data-stu-id="c04ce-143">Before you can use custom properties, you must load them by calling the [loadCustomPropertiesAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) method.</span></span> <span data-ttu-id="c04ce-144">プロパティ バッグを作成したら、[set](/javascript/api/outlook/office.CustomProperties#set-name--value-) と [get](/javascript/api/outlook/office.CustomProperties) メソッドを使用してカスタム プロパティを追加し、取得できます。</span><span class="sxs-lookup"><span data-stu-id="c04ce-144">After you have created the property bag, you can use the [set](/javascript/api/outlook/office.CustomProperties#set-name--value-) and [get](/javascript/api/outlook/office.CustomProperties) methods to add and retrieve custom properties.</span></span> <span data-ttu-id="c04ce-145">プロパティ バッグで行った変更を保存するには、[saveAsync](/javascript/api/outlook/office.CustomProperties#saveasync-callback--asynccontext-) メソッドを使用する必要があります。</span><span class="sxs-lookup"><span data-stu-id="c04ce-145">You must use the [saveAsync](/javascript/api/outlook/office.CustomProperties#saveasync-callback--asynccontext-) method to save any changes that you make to the property bag.</span></span>


 > [!NOTE]
 > <span data-ttu-id="c04ce-146">Outlook on Mac では、カスタム プロパティをキャッシュに入れないため、ユーザーのネットワークがダウンした場合、Outlook on Mac のメール アドインでカスタム プロパティにアクセスできなくなります。</span><span class="sxs-lookup"><span data-stu-id="c04ce-146">Because Outlook on Mac doesn't cache custom properties, if the user's network goes down, mail add-ins in Outlook on Mac would not be able to access their custom properties.</span></span>


### <a name="custom-properties-example"></a><span data-ttu-id="c04ce-147">カスタム プロパティの例</span><span class="sxs-lookup"><span data-stu-id="c04ce-147">Custom properties example</span></span>


<span data-ttu-id="c04ce-p110">以下の例では、カスタム プロパティを使用する単純な Outlook アドインのメソッドのセットを示しています。この例を出発点として、カスタム プロパティを使用するアドインを作成できます。</span><span class="sxs-lookup"><span data-stu-id="c04ce-p110">The following example shows a simplified set of methods for an Outlook add-in that uses custom properties. You can use this example as a starting point for your add-in that uses custom properties.</span></span>

<span data-ttu-id="c04ce-150">以下の例には、次のメソッドが含まれています。</span><span class="sxs-lookup"><span data-stu-id="c04ce-150">This example includes the following methods:</span></span>


- <span data-ttu-id="c04ce-151">[Office.initialize](/javascript/api/office#office-initialize-reason-) -- アドインを初期化し、Exchange Server からカスタム プロパティ バッグを読み込みます。</span><span class="sxs-lookup"><span data-stu-id="c04ce-151">[Office.initialize](/javascript/api/office#office-initialize-reason-) -- Initializes the add-in and loads the custom property bag from the Exchange server.</span></span>

- <span data-ttu-id="c04ce-152">**customPropsCallback** -- サーバーから返されるカスタム プロパティ バッグを取得し、後で使用できるように保存します。</span><span class="sxs-lookup"><span data-stu-id="c04ce-152">**customPropsCallback** -- Gets the custom property bag that is returned from the server and saves it for later use.</span></span>

- <span data-ttu-id="c04ce-153">**updateProperty** -- 特定のプロパティを設定または更新し、変更をサーバーに保存します。</span><span class="sxs-lookup"><span data-stu-id="c04ce-153">**updateProperty** -- Sets or updates a specific property, and then saves the change to the server.</span></span>

- <span data-ttu-id="c04ce-154">**removeProperty** -- プロパティ バッグから特定のプロパティを削除し、その削除をサーバーに保存します。</span><span class="sxs-lookup"><span data-stu-id="c04ce-154">**removeProperty** -- Removes a specific property from the property bag, and then saves the removal to the server.</span></span>


```js
var _mailbox;
var _customProps;

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
  var myProp = _customProps.get("myProp");
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

### <a name="get-custom-properties-using-ews-or-rest"></a><span data-ttu-id="c04ce-155">EWS または REST を使用してカスタム プロパティを取得する</span><span class="sxs-lookup"><span data-stu-id="c04ce-155">Get custom properties using EWS or REST</span></span>

<span data-ttu-id="c04ce-156">EWS または REST を使用して **CustomProperties** を取得する場合は、最初にMAPI ベースの拡張プロパティの名前を決めるようにします。</span><span class="sxs-lookup"><span data-stu-id="c04ce-156">To get **CustomProperties** using EWS or REST, you should first determine the name of its MAPI-based extended property.</span></span> <span data-ttu-id="c04ce-157">その後、MAPI ベースの拡張プロパティを取得するのと同じ方法でそのプロパティを取得できます。</span><span class="sxs-lookup"><span data-stu-id="c04ce-157">You can then get that property in the same way you would get any MAPI-based extended property.</span></span>

#### <a name="how-custom-properties-are-stored-on-an-item"></a><span data-ttu-id="c04ce-158">アイテムでのカスタム プロパティの格納方法</span><span class="sxs-lookup"><span data-stu-id="c04ce-158">How custom properties are stored on an item</span></span>

<span data-ttu-id="c04ce-159">アドインによって設定されたカスタム プロパティは、標準の MAPI ベースのプロパティとは異なります。</span><span class="sxs-lookup"><span data-stu-id="c04ce-159">Custom properties set by an add-in are not equivalent to normal MAPI-based properties.</span></span> <span data-ttu-id="c04ce-160">アドイン api は、 `CustomProperties`すべてのアドインを JSON ペイロードとしてシリアル化した後、名前が1つの MAPI ベースの拡張プロパティ`cecp-<app-guid>`に`<app-guid>`保存されます。このプロパティは、名前が ([ `{00020329-0000-0000-C000-000000000046}`アドインの ID がである) およびプロパティセット GUID です。</span><span class="sxs-lookup"><span data-stu-id="c04ce-160">Add-in APIs serialize all your add-in's `CustomProperties` as a JSON payload and then save them in a single MAPI-based extended property whose name is `cecp-<app-guid>` (`<app-guid>` is your add-in's ID) and property set GUID is `{00020329-0000-0000-C000-000000000046}`.</span></span> <span data-ttu-id="c04ce-161">(このオブジェクトに関する詳細については、「[MS OXCEXT 2.2.5 メール アプリのカスタム プロパティ](https://msdn.microsoft.com/library/hh968549(v=exchg.80).aspx)」を参照してください。) その後、EWS または REST を使用してこの MAPI ベースのプロパティを取得できます。</span><span class="sxs-lookup"><span data-stu-id="c04ce-161">(For more information about this object, see [MS-OXCEXT 2.2.5 Mail App Custom Properties](https://msdn.microsoft.com/library/hh968549(v=exchg.80).aspx).) You can then use EWS or REST to get this MAPI-based property.</span></span>

#### <a name="get-custom-properties-using-ews"></a><span data-ttu-id="c04ce-162">EWS を使用してカスタム プロパティを取得する</span><span class="sxs-lookup"><span data-stu-id="c04ce-162">Get custom properties using EWS</span></span>

<span data-ttu-id="c04ce-163">メールアドインは、EWS の[GetItem](/exchange/client-developer/web-service-reference/getitem-operation)操作`CustomProperties`を使用して MAPI ベースの拡張プロパティを取得できます。</span><span class="sxs-lookup"><span data-stu-id="c04ce-163">Your mail add-in can get the `CustomProperties` MAPI-based extended property by using the EWS [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation.</span></span> <span data-ttu-id="c04ce-164">コール`GetItem`バックトークンを使用して、またはクライアント側で、 [makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)メソッドを使用してサーバー側でアクセスします。</span><span class="sxs-lookup"><span data-stu-id="c04ce-164">Access `GetItem` on the server side by using a callback token, or on the client side by using the [mailbox.makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) method.</span></span> <span data-ttu-id="c04ce-165">`GetItem`要求で、前のセクション`CustomProperties`で説明されている詳細を使用して、プロパティセットに MAPI ベースのプロパティを指定します。このセクションに[は、アイテムにカスタムプロパティが格納](#how-custom-properties-are-stored-on-an-item)されます。</span><span class="sxs-lookup"><span data-stu-id="c04ce-165">In the `GetItem` request, specify the `CustomProperties` MAPI-based property in its property set using the details provided in the preceding section [How custom properties are stored on an item](#how-custom-properties-are-stored-on-an-item).</span></span>

<span data-ttu-id="c04ce-166">次の例では、アイテムとそれのカスタム プロパティを取得する方法を示します。</span><span class="sxs-lookup"><span data-stu-id="c04ce-166">The following example shows how to get an item and its custom properties.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="c04ce-167">次の例では、`<app-guid>` を自分のアドインの ID と置き換えてください。</span><span class="sxs-lookup"><span data-stu-id="c04ce-167">In the following example, replace `<app-guid>` with your add-in's ID.</span></span>

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

<span data-ttu-id="c04ce-168">要求文字列で他のカスタム プロパティを [ExtendedFieldURI](/exchange/client-developer/web-service-reference/extendedfielduri) 要素として指定すると、それらのカスタム プロパティを取得することができます。</span><span class="sxs-lookup"><span data-stu-id="c04ce-168">You can also get more custom properties if you specify them in the request string as other [ExtendedFieldURI](/exchange/client-developer/web-service-reference/extendedfielduri) elements.</span></span>

#### <a name="get-custom-properties-using-rest"></a><span data-ttu-id="c04ce-169">REST を使用してカスタム プロパティを取得する</span><span class="sxs-lookup"><span data-stu-id="c04ce-169">Get custom properties using REST</span></span>

<span data-ttu-id="c04ce-170">アドインでメッセージやイベントに対して REST クエリを作成し、すでにカスタム プロパティを持つメッセージやイベントを取得することができます。</span><span class="sxs-lookup"><span data-stu-id="c04ce-170">In your add-in, you can construct your REST query against messages and events to get the ones that already have custom properties.</span></span> <span data-ttu-id="c04ce-171">「[アイテムでのカスタム プロパティの格納方法](#how-custom-properties-are-stored-on-an-item)」セクションで説明されている詳細を参考にして、クエリに MAPI ベースの拡張プロパティ **CustomProperties** とそのプロパティ セットを含めます。</span><span class="sxs-lookup"><span data-stu-id="c04ce-171">In your query, you should include the **CustomProperties** MAPI-based property and its property set using the details provided in the section [How custom properties are stored on an item](#how-custom-properties-are-stored-on-an-item).</span></span>

<span data-ttu-id="c04ce-172">次の例では、アドインで設定されたいずれかのカスタム プロパティを含むすべてのイベントを取得し、追加のフィルター処理のロジックを適用できるようにプロパティの値を応答に含める方法を示します。</span><span class="sxs-lookup"><span data-stu-id="c04ce-172">The following example shows how to get all events that have any custom properties set by your add-in and ensure that the response includes the value of the property so you can apply further filtering logic.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="c04ce-173">次の例では、`<app-guid>` を自分のアドインの ID と置き換えてください。</span><span class="sxs-lookup"><span data-stu-id="c04ce-173">In the following example, replace `<app-guid>` with your add-in's ID.</span></span>

```rest
GET https://outlook.office.com/api/v2.0/Me/Events?$filter=SingleValueExtendedProperties/Any
  (ep: ep/PropertyId eq 'String {00020329-0000-0000-C000-000000000046}
  Name cecp-<app-guid>' and ep/Value ne null)
  &$expand=SingleValueExtendedProperties($filter=PropertyId eq 'String
  {00020329-0000-0000-C000-000000000046} Name cecp-<app-guid>')
```

<span data-ttu-id="c04ce-174">REST を使用して単一値の MAPI ベースの拡張プロパティを取得するその他の例は、「[singleValueExtendedProperty を取得る](/graph/api/singlevaluelegacyextendedproperty-get?view=graph-rest-1.0)」 を参照してください。</span><span class="sxs-lookup"><span data-stu-id="c04ce-174">For other examples that use REST to get single-value MAPI-based extended properties, see [Get singleValueExtendedProperty](/graph/api/singlevaluelegacyextendedproperty-get?view=graph-rest-1.0).</span></span>

<span data-ttu-id="c04ce-175">次の例では、アイテムとそれのカスタム プロパティを取得する方法を示します。</span><span class="sxs-lookup"><span data-stu-id="c04ce-175">The following example shows how to get an item and its custom properties.</span></span> <span data-ttu-id="c04ce-176">`done` メソッドのコールバック関数では、要求されたカスタム プロパティの一覧は `item.SingleValueExtendedProperties` に含まれます。</span><span class="sxs-lookup"><span data-stu-id="c04ce-176">In the callback function for the `done` method, `item.SingleValueExtendedProperties` contains a list of the requested custom properties.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="c04ce-177">次の例では、`<app-guid>` を自分のアドインの ID と置き換えてください。</span><span class="sxs-lookup"><span data-stu-id="c04ce-177">In the following example, replace `<app-guid>` with your add-in's ID.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="c04ce-178">関連項目</span><span class="sxs-lookup"><span data-stu-id="c04ce-178">See also</span></span>

- [<span data-ttu-id="c04ce-179">MAPI のプロパティの概要</span><span class="sxs-lookup"><span data-stu-id="c04ce-179">MAPI Property Overview</span></span>](/office/client-developer/outlook/mapi/mapi-property-overview)
- [<span data-ttu-id="c04ce-180">Outlook のプロパティの概要</span><span class="sxs-lookup"><span data-stu-id="c04ce-180">Outlook Properties Overview</span></span>](/office/vba/outlook/How-to/Navigation/properties-overview)  
- [<span data-ttu-id="c04ce-181">Outlook アドインからの Outlook REST API の呼び出し</span><span class="sxs-lookup"><span data-stu-id="c04ce-181">Call Outlook REST APIs from an Outlook add-in</span></span>](use-rest-api.md)
- [<span data-ttu-id="c04ce-182">Outlook アドインから Web サービスを呼び出す</span><span class="sxs-lookup"><span data-stu-id="c04ce-182">Call web services from an Outlook add-in</span></span>](web-services.md)
- <span data-ttu-id="c04ce-183">
  [Exchange における EWS のプロパティと拡張プロパティ](/exchange/client-developer/exchange-web-services/properties-and-extended-properties-in-ews-in-exchange)</span><span class="sxs-lookup"><span data-stu-id="c04ce-183">[Properties and extended properties in EWS in Exchange](/exchange/client-developer/exchange-web-services/properties-and-extended-properties-in-ews-in-exchange)</span></span>
- <span data-ttu-id="c04ce-184">
  [Exchange の EWS でのプロパティ セットと応答の図形](/exchange/client-developer/exchange-web-services/property-sets-and-response-shapes-in-ews-in-exchange)</span><span class="sxs-lookup"><span data-stu-id="c04ce-184">[Property sets and response shapes in EWS in Exchange](/exchange/client-developer/exchange-web-services/property-sets-and-response-shapes-in-ews-in-exchange)</span></span>
