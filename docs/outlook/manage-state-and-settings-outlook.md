---
title: アドインの状態と設定Outlook管理する
description: アドインの状態と設定を保持する方法について、Outlookします。
ms.date: 04/29/2021
localization_priority: Normal
ms.openlocfilehash: 6652034ffa6844d22fd725adc5adcc4a4063c1cb
ms.sourcegitcommit: 6057afc1776e1667b231d2e9809d261d372151f6
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/30/2021
ms.locfileid: "52100286"
---
# <a name="manage-state-and-settings-for-an-outlook-add-in"></a><span data-ttu-id="4b4da-103">アドインの状態と設定Outlook管理する</span><span class="sxs-lookup"><span data-stu-id="4b4da-103">Manage state and settings for an Outlook add-in</span></span>

> [!NOTE]
> <span data-ttu-id="4b4da-104">この記事 [を読む](../develop/persisting-add-in-state-and-settings.md) 前に、このドキュメントの **「Core concepts」** セクションの「永続化アドインの状態と設定」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="4b4da-104">Please review [Persisting add-in state and settings](../develop/persisting-add-in-state-and-settings.md) in the **Core concepts** section of this documentation before reading this article.</span></span>

<span data-ttu-id="4b4da-105">Outlookアドインの場合、Office JavaScript API は、次の表に示すとおり、セッション間でアドインの状態を保存する[RoamingSettings](/javascript/api/outlook/office.roamingsettings)オブジェクトと[CustomProperties](/javascript/api/outlook/office.customproperties)オブジェクトを提供します。</span><span class="sxs-lookup"><span data-stu-id="4b4da-105">For Outlook add-ins, the Office JavaScript API provides [RoamingSettings](/javascript/api/outlook/office.roamingsettings) and [CustomProperties](/javascript/api/outlook/office.customproperties) objects for saving add-in state across sessions as described in the following table.</span></span> <span data-ttu-id="4b4da-106">すべてのケースで、保存された設定値は、それを作成したアドインの [Id](../reference/manifest/id.md) にのみ関連付けられます。</span><span class="sxs-lookup"><span data-stu-id="4b4da-106">In all cases, the saved settings values are associated with the [Id](../reference/manifest/id.md) of the add-in that created them.</span></span>

|<span data-ttu-id="4b4da-107">**オブジェクト**</span><span class="sxs-lookup"><span data-stu-id="4b4da-107">**Object**</span></span>|<span data-ttu-id="4b4da-108">**ストレージの場所**</span><span class="sxs-lookup"><span data-stu-id="4b4da-108">**Storage location**</span></span>|
|:-----|:-----|
|[<span data-ttu-id="4b4da-109">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="4b4da-109">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings)|<span data-ttu-id="4b4da-110">アドインがインストールされている、ユーザーの Exchange サーバー メールボックス。</span><span class="sxs-lookup"><span data-stu-id="4b4da-110">The user's Exchange server mailbox where the add-in is installed.</span></span> <span data-ttu-id="4b4da-111">これらの設定はユーザーのサーバー メールボックスに格納されますので、ユーザーと一緒に "ローミング" し、サポートされているクライアントがユーザーのメールボックスにアクセスするコンテキストでアドインを実行するときに使用できます。</span><span class="sxs-lookup"><span data-stu-id="4b4da-111">Because these settings are stored in the user's server mailbox, they can "roam" with the user and are available to the add-in when it is running in the context of any supported client accessing that user's mailbox.</span></span><br/><br/> <span data-ttu-id="4b4da-112">Outlook アドインのローミング設定は、その設定を作成したアドインのみが利用でき、また、アドインがインストールされているメールボックスからのみ利用できます。</span><span class="sxs-lookup"><span data-stu-id="4b4da-112">Outlook add-in roaming settings are available only to the add-in that created them, and only from the mailbox where the add-in is installed.</span></span>|
|[<span data-ttu-id="4b4da-113">CustomProperties</span><span class="sxs-lookup"><span data-stu-id="4b4da-113">CustomProperties</span></span>](/javascript/api/outlook/office.customproperties)|<span data-ttu-id="4b4da-p103">アドインが連携するメッセージ、予定、または会議出席依頼アイテム。 Outlook アドイン アイテムのカスタム プロパティは、そのプロパティを作成したアドインのみが利用でき、また、プロパティが保存されているアイテムからのみ利用できます。</span><span class="sxs-lookup"><span data-stu-id="4b4da-p103">The message, appointment, or meeting request item the add-in is working with. Outlook add-in item custom properties are available only to the add-in that created them, and only from the item where they are saved.</span></span>|

## <a name="how-to-save-settings-in-the-users-mailbox-for-outlook-add-ins-as-roaming-settings"></a><span data-ttu-id="4b4da-116">Outlook アドインでユーザーのメールボックスに設定をローミング設定として保存する方法</span><span class="sxs-lookup"><span data-stu-id="4b4da-116">How to save settings in the user's mailbox for Outlook add-ins as roaming settings</span></span>

<span data-ttu-id="4b4da-117">Outlook アドインは、[RoamingSettings](/javascript/api/outlook/office.roamingsettings) オブジェクトを使用して、ユーザーのメールボックスに固有の、アドインの状態および設定のデータを保存できます。</span><span class="sxs-lookup"><span data-stu-id="4b4da-117">An Outlook add-in can use the [RoamingSettings](/javascript/api/outlook/office.roamingsettings) object to save add-in state and settings data that is specific to the user's mailbox.</span></span> <span data-ttu-id="4b4da-118">このデータには、アドインを実行しているユーザーではなく、Outlook アドインのみがアクセスできます。</span><span class="sxs-lookup"><span data-stu-id="4b4da-118">This data is accessible only by that Outlook add-in on behalf of the user running the add-in.</span></span> <span data-ttu-id="4b4da-119">データはユーザーの Exchange Server メールボックスに格納されます。データには、ユーザーが自分のアカウントにログインして Outlook アドインを実行したときにアクセスできるようになります。</span><span class="sxs-lookup"><span data-stu-id="4b4da-119">The data is stored on the user's Exchange Server mailbox, and is accessible when that user logs into their account and runs the Outlook add-in.</span></span>

### <a name="loading-roaming-settings"></a><span data-ttu-id="4b4da-120">ローミング設定の読み込み</span><span class="sxs-lookup"><span data-stu-id="4b4da-120">Loading roaming settings</span></span>

<span data-ttu-id="4b4da-121">次の JavaScript のコード例は、既存のローミング設定を読み込む方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="4b4da-121">The following JavaScript code example shows how to load existing roaming settings.</span></span>

```js
var _settings = Office.context.roamingSettings;
```

### <a name="creating-or-assigning-a-roaming-setting"></a><span data-ttu-id="4b4da-122">ローミング設定の作成または割り当て</span><span class="sxs-lookup"><span data-stu-id="4b4da-122">Creating or assigning a roaming setting</span></span>

<span data-ttu-id="4b4da-p105">前の例に続けて、次の  `setAppSetting` 関数では、 [RoamingSettings.set](/javascript/api/outlook/office.roamingsettings#set-name--value-) メソッドを使用して、 `cookie` という名前の設定項目に今日の日付を設定、または今日の日付で更新する方法を示しています。次に、 [RoamingSettings.saveAsync](/javascript/api/outlook/office.roamingsettings#saveasync-callback-) メソッドを使用して Exchange Server にすべてのローミング設定を保存し直しています。</span><span class="sxs-lookup"><span data-stu-id="4b4da-p105">Continuing with the preceding example, the following  `setAppSetting` function shows how to use the [RoamingSettings.set](/javascript/api/outlook/office.roamingsettings#set-name--value-) method to set or update a setting named `cookie` with today's date. Then, it saves all the roaming settings back to the Exchange Server with the [RoamingSettings.saveAsync](/javascript/api/outlook/office.roamingsettings#saveasync-callback-) method.</span></span>

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

<span data-ttu-id="4b4da-125">**saveAsync** メソッドは、ローミング設定を非同期で保存し、オプションのコールバック関数を受け取ります。</span><span class="sxs-lookup"><span data-stu-id="4b4da-125">The **saveAsync** method saves roaming settings asynchronously and takes an optional callback function.</span></span> <span data-ttu-id="4b4da-126">このコード例では、`saveMyAppSettingsCallback` という名前のコールバック関数を **saveAsync** メソッドに渡します。</span><span class="sxs-lookup"><span data-stu-id="4b4da-126">This code sample passes a callback function named `saveMyAppSettingsCallback` to the **saveAsync** method.</span></span> <span data-ttu-id="4b4da-127">非同期呼び出しが返されると、`saveMyAppSettingsCallback` 関数の _asyncResult_ パラメーターが [AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトにアクセスします。このオブジェクトを使用すると、**AsyncResult.status** プロパティで操作の成功または失敗を判定することができます。</span><span class="sxs-lookup"><span data-stu-id="4b4da-127">When the asynchronous call returns, the _asyncResult_ parameter of the `saveMyAppSettingsCallback` function provides access to an [AsyncResult](/javascript/api/office/office.asyncresult) object that you can use to determine the success or failure of the operation with the **AsyncResult.status** property.</span></span>

### <a name="removing-a-roaming-setting"></a><span data-ttu-id="4b4da-128">ローミング設定の削除</span><span class="sxs-lookup"><span data-stu-id="4b4da-128">Removing a roaming setting</span></span>

<span data-ttu-id="4b4da-129">また、次の  `removeAppSetting` 関数は、前の例をさらに拡張するものです。この例では、 [RoamingSettings.remove](/javascript/api/outlook/office.roamingsettings#remove-name-) メソッドを使用して `cookie` 設定を削除し、すべてのローミング設定を Exchange Server に保存し直す方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="4b4da-129">Also extending the preceding examples, the following  `removeAppSetting` function, shows how to use the [RoamingSettings.remove](/javascript/api/outlook/office.roamingsettings#remove-name-) method to remove the `cookie` setting and save all the roaming settings back to the Exchange Server.</span></span>

```js
// Remove an application setting.
function removeAppSetting()
{
    _settings.remove("cookie");
    _settings.saveAsync(saveMyAppSettingsCallback);
}
```

## <a name="how-to-save-settings-per-item-for-outlook-add-ins-as-custom-properties"></a><span data-ttu-id="4b4da-130">Outlook アドインでアイテムごとに設定をカスタムプロパティとして保存する方法</span><span class="sxs-lookup"><span data-stu-id="4b4da-130">How to save settings per item for Outlook add-ins as custom properties</span></span>

<span data-ttu-id="4b4da-p107">カスタム プロパティを使用すると、Outlook アドインは処理しているアイテムに関する情報を保存できます。たとえば、Outlook アドインを使用して、メッセージ内の会議の提案から予定を作成する場合は、カスタム プロパティを使用して、会議が作成されたという事実を保存できます。これにより、メッセージを再び開いたときに、Outlook アドインが再び予定の作成を行うことはありません。</span><span class="sxs-lookup"><span data-stu-id="4b4da-p107">Custom properties let your Outlook add-in store information about an item it is working with. For example, if your Outlook add-in creates an appointment from a meeting suggestion in a message, you can use custom properties to store the fact that the meeting was created. This makes sure that if the message is opened again, your Outlook add-in doesn't offer to create the appointment again.</span></span>

<span data-ttu-id="4b4da-p108">メッセージ、予定、または会議出席依頼の特定のアイテムに対してカスタム プロパティを使用するには、その前に、 [Item](/javascript/api/outlook/office.mailbox) オブジェクトの **loadCustomPropertiesAsync** メソッドを呼び出して、プロパティをメモリに読み込む必要があります。現在のアイテムに対してカスタム プロパティが既に設定されている場合は、この時点で Exchange サーバーから読み込まれます。プロパティを読み込んだ後、 [CustomProperties](/javascript/api/outlook/office.customproperties#set-name--value-) オブジェクトの [set](/javascript/api/outlook/office.roamingsettings) メソッドおよび **get** メソッドを使用して、メモリ内のプロパティの追加、更新、および取得を実行できます。アイテムのカスタム プロパティに対して行った変更を保存するには、 [saveAsync](/javascript/api/outlook/office.customproperties#saveasync-callback--asynccontext-) メソッドを使用して、アイテムに加えた変更を Exchange サーバー上で保持する必要があります。</span><span class="sxs-lookup"><span data-stu-id="4b4da-p108">Before you can use custom properties for a particular message, appointment, or meeting request item, you must load the properties into memory by calling the [loadCustomPropertiesAsync](/javascript/api/outlook/office.mailbox) method of the **Item** object. If any custom properties are already set for the current item, they are loaded from the Exchange server at this point. After you have loaded the properties, you can use the [set](/javascript/api/outlook/office.customproperties#set-name--value-) and [get](/javascript/api/outlook/office.roamingsettings) methods of the **CustomProperties** object to add, update, and retrieve properties in memory. To save any changes that you make to the item's custom properties, you must use the [saveAsync](/javascript/api/outlook/office.customproperties#saveasync-callback--asynccontext-) method to persist the changes to the item on the Exchange server.</span></span>

### <a name="custom-properties-example"></a><span data-ttu-id="4b4da-138">カスタム プロパティの例</span><span class="sxs-lookup"><span data-stu-id="4b4da-138">Custom properties example</span></span>

<span data-ttu-id="4b4da-p109">以下の例では、カスタム プロパティを使用する Outlook アドインの一連の関数を、簡略化して示しています。この例を出発点として、カスタム プロパティを使用する Outlook アドインを作成できます。</span><span class="sxs-lookup"><span data-stu-id="4b4da-p109">The following example shows a simplified set of functions for an Outlook add-in that uses custom properties. You can use this example as a starting point for your Outlook add-in that uses custom properties.</span></span>

<span data-ttu-id="4b4da-141">これらの関数を使用する Outlook アドインは、次の例に示すように、`_customProps` 変数で **get** メソッドを呼び出すことによって、任意のカスタム プロパティを取得します。</span><span class="sxs-lookup"><span data-stu-id="4b4da-141">An Outlook add-in that uses these functions retrieves any custom properties by calling the **get** method on the `_customProps` variable, as shown in the following example.</span></span>

```js
var property = _customProps.get("propertyName");
```

<span data-ttu-id="4b4da-142">以下の例には、次の関数が含まれています。</span><span class="sxs-lookup"><span data-stu-id="4b4da-142">This example includes the following functions:</span></span>

|<span data-ttu-id="4b4da-143">**関数名**</span><span class="sxs-lookup"><span data-stu-id="4b4da-143">**Function name**</span></span>|<span data-ttu-id="4b4da-144">**説明**</span><span class="sxs-lookup"><span data-stu-id="4b4da-144">**Description**</span></span>|
|:-----|:-----|
| `Office.initialize`|<span data-ttu-id="4b4da-145">アドインを初期化し、Exchange サーバーから現在のアイテムのカスタム プロパティを読み込みます。</span><span class="sxs-lookup"><span data-stu-id="4b4da-145">Initializes the add-in and loads the custom properties for the current item from the Exchange server.</span></span>|
| `customPropsCallback`|<span data-ttu-id="4b4da-146">Exchange サーバーから返されるカスタム プロパティを取得し、後で使用できるように保存します。</span><span class="sxs-lookup"><span data-stu-id="4b4da-146">Gets the custom properties that are returned from the Exchange server and saves it for later use.</span></span>|
| `updateProperty`|<span data-ttu-id="4b4da-147">特定のプロパティを設定または更新し、その変更を Exchange サーバーに保存します。</span><span class="sxs-lookup"><span data-stu-id="4b4da-147">Sets or updates a specific property, and then saves the change to the Exchange server.</span></span>|
| `removeProperty`|<span data-ttu-id="4b4da-148">特定のプロパティを削除し、その削除を Exchange サーバーに保存します。</span><span class="sxs-lookup"><span data-stu-id="4b4da-148">Removes a specific property, and then persists the removal to the Exchange server.</span></span>|
| `saveCallback`|<span data-ttu-id="4b4da-149">`updateProperty` 関数および `removeProperty` 関数内で **saveAsync** メソッドを呼び出すためのコールバック。</span><span class="sxs-lookup"><span data-stu-id="4b4da-149">Callback for calls to the **saveAsync** method in the `updateProperty` and `removeProperty` functions.</span></span>|

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

### <a name="platform-behavior-in-emails"></a><span data-ttu-id="4b4da-150">電子メールでのプラットフォームの動作</span><span class="sxs-lookup"><span data-stu-id="4b4da-150">Platform behavior in emails</span></span>

<span data-ttu-id="4b4da-151">次の表に、さまざまなクライアントのメールに保存されたカスタム プロパティの動作Outlook示します。</span><span class="sxs-lookup"><span data-stu-id="4b4da-151">The following table summarizes saved custom properties behavior in emails for various Outlook clients.</span></span>

|<span data-ttu-id="4b4da-152">シナリオ</span><span class="sxs-lookup"><span data-stu-id="4b4da-152">Scenario</span></span>|<span data-ttu-id="4b4da-153">Windows</span><span class="sxs-lookup"><span data-stu-id="4b4da-153">Windows</span></span>|<span data-ttu-id="4b4da-154">Web</span><span class="sxs-lookup"><span data-stu-id="4b4da-154">Web</span></span>|<span data-ttu-id="4b4da-155">Mac</span><span class="sxs-lookup"><span data-stu-id="4b4da-155">Mac</span></span>|
|---|---|---|---|
|<span data-ttu-id="4b4da-156">新規作成</span><span class="sxs-lookup"><span data-stu-id="4b4da-156">New compose</span></span>|<span data-ttu-id="4b4da-157">null</span><span class="sxs-lookup"><span data-stu-id="4b4da-157">null</span></span>|<span data-ttu-id="4b4da-158">null</span><span class="sxs-lookup"><span data-stu-id="4b4da-158">null</span></span>|<span data-ttu-id="4b4da-159">null</span><span class="sxs-lookup"><span data-stu-id="4b4da-159">null</span></span>|
|<span data-ttu-id="4b4da-160">返信、すべて返信</span><span class="sxs-lookup"><span data-stu-id="4b4da-160">Reply, reply all</span></span>|<span data-ttu-id="4b4da-161">null</span><span class="sxs-lookup"><span data-stu-id="4b4da-161">null</span></span>|<span data-ttu-id="4b4da-162">null</span><span class="sxs-lookup"><span data-stu-id="4b4da-162">null</span></span>|<span data-ttu-id="4b4da-163">null</span><span class="sxs-lookup"><span data-stu-id="4b4da-163">null</span></span>|
|<span data-ttu-id="4b4da-164">転送</span><span class="sxs-lookup"><span data-stu-id="4b4da-164">Forward</span></span>|<span data-ttu-id="4b4da-165">親のプロパティを読み込む</span><span class="sxs-lookup"><span data-stu-id="4b4da-165">Loads parent's properties</span></span>|<span data-ttu-id="4b4da-166">null</span><span class="sxs-lookup"><span data-stu-id="4b4da-166">null</span></span>|<span data-ttu-id="4b4da-167">null</span><span class="sxs-lookup"><span data-stu-id="4b4da-167">null</span></span>|
|<span data-ttu-id="4b4da-168">新しい作成から送信されたアイテム</span><span class="sxs-lookup"><span data-stu-id="4b4da-168">Sent item from new compose</span></span>|<span data-ttu-id="4b4da-169">null</span><span class="sxs-lookup"><span data-stu-id="4b4da-169">null</span></span>|<span data-ttu-id="4b4da-170">null</span><span class="sxs-lookup"><span data-stu-id="4b4da-170">null</span></span>|<span data-ttu-id="4b4da-171">null</span><span class="sxs-lookup"><span data-stu-id="4b4da-171">null</span></span>|
|<span data-ttu-id="4b4da-172">返信または返信から送信されたアイテムすべて</span><span class="sxs-lookup"><span data-stu-id="4b4da-172">Sent item from reply or reply all</span></span>|<span data-ttu-id="4b4da-173">null</span><span class="sxs-lookup"><span data-stu-id="4b4da-173">null</span></span>|<span data-ttu-id="4b4da-174">null</span><span class="sxs-lookup"><span data-stu-id="4b4da-174">null</span></span>|<span data-ttu-id="4b4da-175">null</span><span class="sxs-lookup"><span data-stu-id="4b4da-175">null</span></span>|
|<span data-ttu-id="4b4da-176">転送から送信されたアイテム</span><span class="sxs-lookup"><span data-stu-id="4b4da-176">Sent item from forward</span></span>|<span data-ttu-id="4b4da-177">保存されていない場合、親のプロパティを削除します</span><span class="sxs-lookup"><span data-stu-id="4b4da-177">Removes parent's properties if not saved</span></span>|<span data-ttu-id="4b4da-178">null</span><span class="sxs-lookup"><span data-stu-id="4b4da-178">null</span></span>|<span data-ttu-id="4b4da-179">null</span><span class="sxs-lookup"><span data-stu-id="4b4da-179">null</span></span>|

<span data-ttu-id="4b4da-180">次の操作で状況を処理Windows。</span><span class="sxs-lookup"><span data-stu-id="4b4da-180">To handle the situation on Windows:</span></span>

1. <span data-ttu-id="4b4da-181">アドインの初期化時に既存のプロパティを確認し、それらを保持するか、必要に応じてオフにしてください。</span><span class="sxs-lookup"><span data-stu-id="4b4da-181">Check for existing properties on initializing your add-in, and keep them or clear them as needed.</span></span>
1. <span data-ttu-id="4b4da-182">カスタム プロパティを設定する場合は、メッセージの読み取り中またはアドインの読み取りモードでカスタム プロパティが追加されたかどうかを示す追加のプロパティを含める必要があります。</span><span class="sxs-lookup"><span data-stu-id="4b4da-182">When setting custom properties, include an additional property to indicate whether the custom properties were added during message read or by Read mode of the add-in.</span></span> <span data-ttu-id="4b4da-183">これは、プロパティが作成中に作成されたのか、親から継承されたのかを区別するのに役立ちます。</span><span class="sxs-lookup"><span data-stu-id="4b4da-183">This will help you differentiate if the property was created during compose or inherited from the parent.</span></span>
1. <span data-ttu-id="4b4da-184">[item.getComposeTypeAsync](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview&preserve-view=true#getComposeTypeAsync_options__callback_) (現在プレビュー中) を使用して、ユーザーが電子メールまたは返信を転送している場合も確認できます。</span><span class="sxs-lookup"><span data-stu-id="4b4da-184">You can also use [item.getComposeTypeAsync](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview&preserve-view=true#getComposeTypeAsync_options__callback_) (currently in preview) to check if the user is forwarding an email or replying.</span></span>

## <a name="see-also"></a><span data-ttu-id="4b4da-185">関連項目</span><span class="sxs-lookup"><span data-stu-id="4b4da-185">See also</span></span>

- [<span data-ttu-id="4b4da-186">アドインの状態および設定を保持する</span><span class="sxs-lookup"><span data-stu-id="4b4da-186">Persisting add-in state and settings</span></span>](../develop/persisting-add-in-state-and-settings.md)
- [<span data-ttu-id="4b4da-187">Office アドインを初期化する</span><span class="sxs-lookup"><span data-stu-id="4b4da-187">Initialize your Office Add-in</span></span>](../develop/initialize-add-in.md)
