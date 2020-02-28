---
title: Outlook アドインの状態と設定を管理する
description: Outlook アドインのアドインの状態と設定を保存する方法について説明します。
ms.date: 02/27/2020
localization_priority: Normal
ms.openlocfilehash: 7a76da625faab98de1f6ef6d32e0274056dba9f2
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/27/2020
ms.locfileid: "42325348"
---
# <a name="manage-state-and-settings-for-an-outlook-add-in"></a><span data-ttu-id="6db6a-103">Outlook アドインの状態と設定を管理する</span><span class="sxs-lookup"><span data-stu-id="6db6a-103">Manage state and settings for an Outlook add-in</span></span>

> [!NOTE]
> <span data-ttu-id="6db6a-104">この記事を読む前に、このドキュメントの「**コア概念**」セクションの「[アドインの状態と設定を保持](../develop/persisting-add-in-state-and-settings.md)する」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="6db6a-104">Please review [Persisting add-in state and settings](../develop/persisting-add-in-state-and-settings.md) in the **Core concepts** section of this documentation before reading this article.</span></span>

<span data-ttu-id="6db6a-105">Outlook アドインの場合、Office JavaScript API は、次の表に示すように、セッション間でアドインの状態を保存するための[RoamingSettings](/javascript/api/outlook/office.roamingsettings)オブジェクトと[CustomProperties](/javascript/api/outlook/office.customproperties)オブジェクトを提供します。</span><span class="sxs-lookup"><span data-stu-id="6db6a-105">For Outlook add-ins, the Office JavaScript API provides [RoamingSettings](/javascript/api/outlook/office.roamingsettings) and [CustomProperties](/javascript/api/outlook/office.customproperties) objects for saving add-in state across sessions as described in the following table.</span></span> <span data-ttu-id="6db6a-106">すべてのケースで、保存された設定値は、それを作成したアドインの [Id](/office/dev/add-ins/reference/manifest/id) にのみ関連付けられます。</span><span class="sxs-lookup"><span data-stu-id="6db6a-106">In all cases, the saved settings values are associated with the [Id](/office/dev/add-ins/reference/manifest/id) of the add-in that created them.</span></span>

|<span data-ttu-id="6db6a-107">**オブジェクト**</span><span class="sxs-lookup"><span data-stu-id="6db6a-107">**Object**</span></span>|<span data-ttu-id="6db6a-108">**ストレージの場所**</span><span class="sxs-lookup"><span data-stu-id="6db6a-108">**Storage location**</span></span>|
|:-----|:-----|:-----|
|[<span data-ttu-id="6db6a-109">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="6db6a-109">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings)|<span data-ttu-id="6db6a-110">アドインがインストールされている、ユーザーの Exchange サーバー メールボックス。</span><span class="sxs-lookup"><span data-stu-id="6db6a-110">The user's Exchange server mailbox where the add-in is installed.</span></span> <span data-ttu-id="6db6a-111">これらの設定はユーザーのサーバー メールボックスに保存されるので、ユーザーと共に "ローミング" でき、そのユーザーのメールボックスにアクセスしている、サポートされているクライアント ホスト アプリケーションまたはブラウザーのコンテキストでアドインが実行されている場合、そのアドインでこれらの設定を利用できます。</span><span class="sxs-lookup"><span data-stu-id="6db6a-111">Because these settings are stored in the user's server mailbox, they can "roam" with the user and are available to the add-in when it is running in the context of any supported client host application or browser accessing that user's mailbox.</span></span><br/><br/> <span data-ttu-id="6db6a-112">Outlook アドインのローミング設定は、その設定を作成したアドインのみが利用でき、また、アドインがインストールされているメールボックスからのみ利用できます。</span><span class="sxs-lookup"><span data-stu-id="6db6a-112">Outlook add-in roaming settings are available only to the add-in that created them, and only from the mailbox where the add-in is installed.</span></span>|
|[<span data-ttu-id="6db6a-113">CustomProperties</span><span class="sxs-lookup"><span data-stu-id="6db6a-113">CustomProperties</span></span>](/javascript/api/outlook/office.customproperties)|<span data-ttu-id="6db6a-p103">アドインが連携するメッセージ、予定、または会議出席依頼アイテム。 Outlook アドイン アイテムのカスタム プロパティは、そのプロパティを作成したアドインのみが利用でき、また、プロパティが保存されているアイテムからのみ利用できます。</span><span class="sxs-lookup"><span data-stu-id="6db6a-p103">The message, appointment, or meeting request item the add-in is working with. Outlook add-in item custom properties are available only to the add-in that created them, and only from the item where they are saved.</span></span>|

## <a name="how-to-save-settings-in-the-users-mailbox-for-outlook-add-ins-as-roaming-settings"></a><span data-ttu-id="6db6a-116">Outlook アドインでユーザーのメールボックスに設定をローミング設定として保存する方法</span><span class="sxs-lookup"><span data-stu-id="6db6a-116">How to save settings in the user's mailbox for Outlook add-ins as roaming settings</span></span>

<span data-ttu-id="6db6a-117">Outlook アドインは、[RoamingSettings](/javascript/api/outlook/office.roamingsettings) オブジェクトを使用して、ユーザーのメールボックスに固有の、アドインの状態および設定のデータを保存できます。</span><span class="sxs-lookup"><span data-stu-id="6db6a-117">An Outlook add-in can use the [RoamingSettings](/javascript/api/outlook/office.roamingsettings) object to save add-in state and settings data that is specific to the user's mailbox.</span></span> <span data-ttu-id="6db6a-118">このデータには、アドインを実行しているユーザーではなく、Outlook アドインのみがアクセスできます。</span><span class="sxs-lookup"><span data-stu-id="6db6a-118">This data is accessible only by that Outlook add-in on behalf of the user running the add-in.</span></span> <span data-ttu-id="6db6a-119">データはユーザーの Exchange Server メールボックスに格納されます。データには、ユーザーが自分のアカウントにログインして Outlook アドインを実行したときにアクセスできるようになります。</span><span class="sxs-lookup"><span data-stu-id="6db6a-119">The data is stored on the user's Exchange Server mailbox, and is accessible when that user logs into their account and runs the Outlook add-in.</span></span>

### <a name="loading-roaming-settings"></a><span data-ttu-id="6db6a-120">ローミング設定の読み込み</span><span class="sxs-lookup"><span data-stu-id="6db6a-120">Loading roaming settings</span></span>

<span data-ttu-id="6db6a-p105">通常、Outlook アドインでは、 [Office.initialize](/javascript/api/office) イベント ハンドラーでローミング設定を読み込みます。次の JavaScript のコード例は、既存のローミング設定を読み込む方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="6db6a-p105">An Outlook add-in typically loads roaming settings in the [Office.initialize](/javascript/api/office) event handler. The following JavaScript code example shows how to load existing roaming settings.</span></span>

```js
var _mailbox;
var _settings;

// The initialize function is required for all add-ins.
Office.initialize = function (reason) {
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
    // After the DOM is loaded, add-in-specific code can run.
   // Initialize instance variables to access API objects.
    _mailbox = Office.context.mailbox;
    _settings = Office.context.roamingSettings;
    });
}
```

### <a name="creating-or-assigning-a-roaming-setting"></a><span data-ttu-id="6db6a-123">ローミング設定の作成または割り当て</span><span class="sxs-lookup"><span data-stu-id="6db6a-123">Creating or assigning a roaming setting</span></span>

<span data-ttu-id="6db6a-p106">前の例に続けて、次の  `setAppSetting` 関数では、 [RoamingSettings.set](/javascript/api/outlook/office.roamingsettings#set-name--value-) メソッドを使用して、 `cookie` という名前の設定項目に今日の日付を設定、または今日の日付で更新する方法を示しています。次に、 [RoamingSettings.saveAsync](/javascript/api/outlook/office.roamingsettings#saveasync-callback-) メソッドを使用して Exchange Server にすべてのローミング設定を保存し直しています。</span><span class="sxs-lookup"><span data-stu-id="6db6a-p106">Continuing with the preceding example, the following  `setAppSetting` function shows how to use the [RoamingSettings.set](/javascript/api/outlook/office.roamingsettings#set-name--value-) method to set or update a setting named `cookie` with today's date. Then, it saves all the roaming settings back to the Exchange Server with the [RoamingSettings.saveAsync](/javascript/api/outlook/office.roamingsettings#saveasync-callback-) method.</span></span>

```js
// Set an add-in setting.
function setAppSetting() {
    _settings.set("cookie", Date());
    _settings.saveAsync(saveMyAppSettingsCallback);
}

// Saves all roaming settings.
function saveMyAppSettingsCallback(asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        // Handle the failure.
    }
}
```

<span data-ttu-id="6db6a-126">**saveAsync** メソッドは、ローミング設定を非同期で保存し、オプションのコールバック関数を受け取ります。</span><span class="sxs-lookup"><span data-stu-id="6db6a-126">The **saveAsync** method saves roaming settings asynchronously and takes an optional callback function.</span></span> <span data-ttu-id="6db6a-127">このコード例では、`saveMyAppSettingsCallback` という名前のコールバック関数を **saveAsync** メソッドに渡します。</span><span class="sxs-lookup"><span data-stu-id="6db6a-127">This code sample passes a callback function named `saveMyAppSettingsCallback` to the **saveAsync** method.</span></span> <span data-ttu-id="6db6a-128">非同期呼び出しが返されると、`saveMyAppSettingsCallback` 関数の _asyncResult_ パラメーターが [AsyncResult](/javascript/api/outlook) オブジェクトにアクセスします。このオブジェクトを使用すると、**AsyncResult.status** プロパティで操作の成功または失敗を判定することができます。</span><span class="sxs-lookup"><span data-stu-id="6db6a-128">When the asynchronous call returns, the _asyncResult_ parameter of the `saveMyAppSettingsCallback` function provides access to an [AsyncResult](/javascript/api/outlook) object that you can use to determine the success or failure of the operation with the **AsyncResult.status** property.</span></span>

### <a name="removing-a-roaming-setting"></a><span data-ttu-id="6db6a-129">ローミング設定の削除</span><span class="sxs-lookup"><span data-stu-id="6db6a-129">Removing a roaming setting</span></span>

<span data-ttu-id="6db6a-130">また、次の  `removeAppSetting` 関数は、前の例をさらに拡張するものです。この例では、 [RoamingSettings.remove](/javascript/api/outlook/office.roamingsettings#remove-name-) メソッドを使用して `cookie` 設定を削除し、すべてのローミング設定を Exchange Server に保存し直す方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="6db6a-130">Also extending the preceding examples, the following  `removeAppSetting` function, shows how to use the [RoamingSettings.remove](/javascript/api/outlook/office.roamingsettings#remove-name-) method to remove the `cookie` setting and save all the roaming settings back to the Exchange Server.</span></span>

```js
// Remove an application setting.
function removeAppSetting()
{
    _settings.remove("cookie");
    _settings.saveAsync(saveMyAppSettingsCallback);
}
```

## <a name="how-to-save-settings-per-item-for-outlook-add-ins-as-custom-properties"></a><span data-ttu-id="6db6a-131">Outlook アドインでアイテムごとに設定をカスタムプロパティとして保存する方法</span><span class="sxs-lookup"><span data-stu-id="6db6a-131">How to save settings per item for Outlook add-ins as custom properties</span></span>

<span data-ttu-id="6db6a-p108">カスタム プロパティを使用すると、Outlook アドインは処理しているアイテムに関する情報を保存できます。たとえば、Outlook アドインを使用して、メッセージ内の会議の提案から予定を作成する場合は、カスタム プロパティを使用して、会議が作成されたという事実を保存できます。これにより、メッセージを再び開いたときに、Outlook アドインが再び予定の作成を行うことはありません。</span><span class="sxs-lookup"><span data-stu-id="6db6a-p108">Custom properties let your Outlook add-in store information about an item it is working with. For example, if your Outlook add-in creates an appointment from a meeting suggestion in a message, you can use custom properties to store the fact that the meeting was created. This makes sure that if the message is opened again, your Outlook add-in doesn't offer to create the appointment again.</span></span>

<span data-ttu-id="6db6a-p109">メッセージ、予定、または会議出席依頼の特定のアイテムに対してカスタム プロパティを使用するには、その前に、 [Item](/javascript/api/outlook/office.mailbox) オブジェクトの **loadCustomPropertiesAsync** メソッドを呼び出して、プロパティをメモリに読み込む必要があります。現在のアイテムに対してカスタム プロパティが既に設定されている場合は、この時点で Exchange サーバーから読み込まれます。プロパティを読み込んだ後、 [CustomProperties](/javascript/api/outlook/office.customproperties#set-name--value-) オブジェクトの [set](/javascript/api/outlook/office.roamingsettings) メソッドおよび **get** メソッドを使用して、メモリ内のプロパティの追加、更新、および取得を実行できます。アイテムのカスタム プロパティに対して行った変更を保存するには、 [saveAsync](/javascript/api/outlook/office.customproperties#saveasync-callback--asynccontext-) メソッドを使用して、アイテムに加えた変更を Exchange サーバー上で保持する必要があります。</span><span class="sxs-lookup"><span data-stu-id="6db6a-p109">Before you can use custom properties for a particular message, appointment, or meeting request item, you must load the properties into memory by calling the [loadCustomPropertiesAsync](/javascript/api/outlook/office.mailbox) method of the **Item** object. If any custom properties are already set for the current item, they are loaded from the Exchange server at this point. After you have loaded the properties, you can use the [set](/javascript/api/outlook/office.customproperties#set-name--value-) and [get](/javascript/api/outlook/office.roamingsettings) methods of the **CustomProperties** object to add, update, and retrieve properties in memory. To save any changes that you make to the item's custom properties, you must use the [saveAsync](/javascript/api/outlook/office.customproperties#saveasync-callback--asynccontext-) method to persist the changes to the item on the Exchange server.</span></span>

### <a name="custom-properties-example"></a><span data-ttu-id="6db6a-139">カスタム プロパティの例</span><span class="sxs-lookup"><span data-stu-id="6db6a-139">Custom properties example</span></span>

<span data-ttu-id="6db6a-p110">以下の例では、カスタム プロパティを使用する Outlook アドインの一連の関数を、簡略化して示しています。この例を出発点として、カスタム プロパティを使用する Outlook アドインを作成できます。</span><span class="sxs-lookup"><span data-stu-id="6db6a-p110">The following example shows a simplified set of functions for an Outlook add-in that uses custom properties. You can use this example as a starting point for your Outlook add-in that uses custom properties.</span></span> 

<span data-ttu-id="6db6a-142">これらの関数を使用する Outlook アドインは、次の例に示すように、`_customProps` 変数で **get** メソッドを呼び出すことによって、任意のカスタム プロパティを取得します。</span><span class="sxs-lookup"><span data-stu-id="6db6a-142">An Outlook add-in that uses these functions retrieves any custom properties by calling the **get** method on the `_customProps` variable, as shown in the following example.</span></span>

```js
var property = _customProps.get("propertyName");
```

<span data-ttu-id="6db6a-143">以下の例には、次の関数が含まれています。</span><span class="sxs-lookup"><span data-stu-id="6db6a-143">This example includes the following functions:</span></span>

|<span data-ttu-id="6db6a-144">**関数名**</span><span class="sxs-lookup"><span data-stu-id="6db6a-144">**Function name**</span></span>|<span data-ttu-id="6db6a-145">**説明**</span><span class="sxs-lookup"><span data-stu-id="6db6a-145">**Description**</span></span>|
|:-----|:-----|
| `Office.initialize`|<span data-ttu-id="6db6a-146">アドインを初期化し、Exchange サーバーから現在のアイテムのカスタム プロパティを読み込みます。</span><span class="sxs-lookup"><span data-stu-id="6db6a-146">Initializes the add-in and loads the custom properties for the current item from the Exchange server.</span></span>|
| `customPropsCallback`|<span data-ttu-id="6db6a-147">Exchange サーバーから返されるカスタム プロパティを取得し、後で使用できるように保存します。</span><span class="sxs-lookup"><span data-stu-id="6db6a-147">Gets the custom properties that are returned from the Exchange server and saves it for later use.</span></span>|
| `updateProperty`|<span data-ttu-id="6db6a-148">特定のプロパティを設定または更新し、その変更を Exchange サーバーに保存します。</span><span class="sxs-lookup"><span data-stu-id="6db6a-148">Sets or updates a specific property, and then saves the change to the Exchange server.</span></span>|
| `removeProperty`|<span data-ttu-id="6db6a-149">特定のプロパティを削除し、その削除を Exchange サーバーに保存します。</span><span class="sxs-lookup"><span data-stu-id="6db6a-149">Removes a specific property, and then persists the removal to the Exchange server.</span></span>|
| `saveCallback`|<span data-ttu-id="6db6a-150">`updateProperty` 関数および `removeProperty` 関数内で **saveAsync** メソッドを呼び出すためのコールバック。</span><span class="sxs-lookup"><span data-stu-id="6db6a-150">Callback for calls to the **saveAsync** method in the `updateProperty` and `removeProperty` functions.</span></span>|

```js
var _mailbox;
var _customProps;

// The initialize function is required for all add-ins.
Office.initialize = function (reason) {
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
    // After the DOM is loaded, add-in-specific code can run.
    _mailbox = Office.context.mailbox;
    _mailbox.item.loadCustomPropertiesAsync(customPropsCallback);
    });
}

// Get the item's custom properties from the server and save for later use.
function customPropsCallback(asyncResult) {
    _customProps = asyncResult.value;
}

// Sets or updates the specified property, and then saves the change
// to the server.
function updateProperty(name, value) {
    _customProps.set(name, value);
    _customProps.saveAsync(saveCallback);
}

// Removes the specified property, and then persists the removal
// to the server.
function removeProperty(name) {
   _customProps.remove(name);
   _customProps.saveAsync(saveCallback);
}

// Callback for calls to saveAsync method.
function saveCallback(asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        // Handle the failure.
    }
}
```

## <a name="see-also"></a><span data-ttu-id="6db6a-151">関連項目</span><span class="sxs-lookup"><span data-stu-id="6db6a-151">See also</span></span>

- [<span data-ttu-id="6db6a-152">アドインの状態および設定を保持する</span><span class="sxs-lookup"><span data-stu-id="6db6a-152">Persisting add-in state and settings</span></span>](../develop/persisting-add-in-state-and-settings.md)
- [<span data-ttu-id="6db6a-153">Office アドインを初期化する</span><span class="sxs-lookup"><span data-stu-id="6db6a-153">Initialize your Office Add-in</span></span>](../develop/initialize-add-in.md)