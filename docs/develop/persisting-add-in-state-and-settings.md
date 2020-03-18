---
title: アドインの状態および設定を保持する
description: ブラウザーコントロールのステートレス環境で実行されている Office アドイン web アプリケーションでデータを永続化する方法について説明します。
ms.date: 02/27/2020
localization_priority: Normal
ms.openlocfilehash: 7d66a8693c18dbc7f2be59b2799db7429681a57f
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/17/2020
ms.locfileid: "42719393"
---
# <a name="persisting-add-in-state-and-settings"></a><span data-ttu-id="02c83-103">アドインの状態および設定を保持する</span><span class="sxs-lookup"><span data-stu-id="02c83-103">Persisting add-in state and settings</span></span>

[!include[information about the common API](../includes/alert-common-api-info.md)]

<span data-ttu-id="02c83-p101">Office アドインは、基本的にブラウザー コントロールのステートレス環境で動作する Web アプリケーションです。したがって、アドインでは、そのアドインを使用するセッション間で特定の操作または機能を継続して維持するためのデータを保持することが必要な場合があります。たとえば、アドインには、ユーザーの優先ビューや既定の場所など、アドインで保存しておき、アドインが次回初期化されたときにリロードする必要があるカスタム設定または他の値が含まれる場合があります。その場合は、次のようにします。</span><span class="sxs-lookup"><span data-stu-id="02c83-p101">Office Add-ins are essentially web applications running in the stateless environment of a browser control. As a result, your add-in may need to persist data to maintain the continuity of certain operations or features across sessions of using your add-in. For example, your add-in may have custom settings or other values that it needs to save and reload the next time it's initialized, such as a user's preferred view or default location. To do that, you can:</span></span>

- <span data-ttu-id="02c83-108">次のいずれかの方法でデータを格納する Office JavaScript API のメンバーを使用します。</span><span class="sxs-lookup"><span data-stu-id="02c83-108">Use members of the Office JavaScript API that store data as either:</span></span>
    -  <span data-ttu-id="02c83-109">アドインの種類応じた場所に保存されるプロパティ バッグ内の名前と値の組。</span><span class="sxs-lookup"><span data-stu-id="02c83-109">Name/value pairs in a property bag stored in a location that depends on add-in type.</span></span>
    -  <span data-ttu-id="02c83-110">ドキュメント内に保存されるカスタム XML。</span><span class="sxs-lookup"><span data-stu-id="02c83-110">Custom XML stored in the document.</span></span>

- <span data-ttu-id="02c83-111">基になるブラウザー コントロールによって提供される技術である、ブラウザーの Cookie、または HTML5 Web ストレージ ([localStorage](https://developer.mozilla.org/docs/Web/API/Window/localStorage) または [sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage)) を使用します。</span><span class="sxs-lookup"><span data-stu-id="02c83-111">Use techniques provided by the underlying browser control: browser cookies, or HTML5 web storage ([localStorage](https://developer.mozilla.org/docs/Web/API/Window/localStorage) or [sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage)).</span></span>

<span data-ttu-id="02c83-112">この記事では、Office JavaScript API を使用してアドインの状態を保持する方法に焦点を当てます。</span><span class="sxs-lookup"><span data-stu-id="02c83-112">This article focuses on how to use the Office JavaScript API to persist add-in state.</span></span> <span data-ttu-id="02c83-113">ブラウザーの Cookie および Web ストレージの使用例については、「 [Excel-Add-in-JavaScript-PersistCustomSettings](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="02c83-113">For examples of using browser cookies and web storage, see the [Excel-Add-in-JavaScript-PersistCustomSettings](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings).</span></span>

## <a name="persisting-add-in-state-and-settings-with-the-office-javascript-api"></a><span data-ttu-id="02c83-114">Office JavaScript API を使用してアドインの状態と設定を保持する</span><span class="sxs-lookup"><span data-stu-id="02c83-114">Persisting add-in state and settings with the Office JavaScript API</span></span>

<span data-ttu-id="02c83-115">Office JavaScript API には、次の表に示すように、セッション間でアドインの状態を保存するための[設定](/javascript/api/office/office.settings)、 [RoamingSettings](/javascript/api/outlook/office.roamingsettings)、および[CustomProperties](/javascript/api/outlook/office.customproperties)オブジェクトが用意されています。</span><span class="sxs-lookup"><span data-stu-id="02c83-115">The Office JavaScript API provides the [Settings](/javascript/api/office/office.settings), [RoamingSettings](/javascript/api/outlook/office.roamingsettings), and [CustomProperties](/javascript/api/outlook/office.customproperties) objects for saving add-in state across sessions as described in the following table.</span></span> <span data-ttu-id="02c83-116">すべてのケースで、保存された設定値は、それを作成したアドインの [Id](../reference/manifest/id.md) にのみ関連付けられます。</span><span class="sxs-lookup"><span data-stu-id="02c83-116">In all cases, the saved settings values are associated with the [Id](../reference/manifest/id.md) of the add-in that created them.</span></span>

|<span data-ttu-id="02c83-117">**オブジェクト**</span><span class="sxs-lookup"><span data-stu-id="02c83-117">**Object**</span></span>|<span data-ttu-id="02c83-118">**アドインの種類のサポート**</span><span class="sxs-lookup"><span data-stu-id="02c83-118">**Add-in type support**</span></span>|<span data-ttu-id="02c83-119">**ストレージの場所**</span><span class="sxs-lookup"><span data-stu-id="02c83-119">**Storage location**</span></span>|<span data-ttu-id="02c83-120">**サポートされる Office のホスト**</span><span class="sxs-lookup"><span data-stu-id="02c83-120">**Office host support**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="02c83-121">Settings</span><span class="sxs-lookup"><span data-stu-id="02c83-121">Settings</span></span>](/javascript/api/office/office.settings)|<span data-ttu-id="02c83-122">コンテンツおよび作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="02c83-122">content and task pane</span></span>|<span data-ttu-id="02c83-123">アドインが連携しているドキュメント、スプレッドシート、またはプレゼンテーション。コンテンツおよび作業ウィンドウのアドイン設定は、その設定が保存されているドキュメントから、その設定を作成したアドインで使用できます。</span><span class="sxs-lookup"><span data-stu-id="02c83-123">The document, spreadsheet, or presentation the add-in is working with.Content and task pane add-in settings are available to the add-in that created them from the document where they are saved.</span></span><br/><br/><span data-ttu-id="02c83-p104">**重要:\*\*\*\*Settings** オブジェクトを使用して、パスワードおよびその他の機密の個人を特定できる情報 (PII) を保存しないでください。保存されたデータはユーザーに対して表示されませんが、ドキュメントの一部として保存されているため、ドキュメントのファイル形式を直接読み取ることでアクセスできます。アドインによる PII の使用と、アドインが必要とするすべての PII の保存は、開発するアドインをユーザーのセキュリティが保護されるリソースとしてホストするサーバーのみで行うよう制限する必要があります。</span><span class="sxs-lookup"><span data-stu-id="02c83-p104">**Important:** Don't store passwords and other sensitive personally identifiable information (PII) with the **Settings** object. The data saved isn't visible to end users, but it is stored as part of the document, which is accessible by reading the document's file format directly. You should limit your add-in's use of PII and store any PII required by your add-in only on the server hosting your add-in as a user-secured resource.</span></span>|<span data-ttu-id="02c83-127">Word、Excel、または PowerPoint</span><span class="sxs-lookup"><span data-stu-id="02c83-127">Word, Excel, or PowerPoint</span></span><br/><br/> <span data-ttu-id="02c83-p105">**メモ:** Project 2013 の作業ウィンドウ アドインでは、アドインの状態または設定を保存するための **Settings** API をサポートしていません。ただし、Project (および他の Office ホスト アプリケーション) で動作するアドインの場合は、ブラウザーの Cookie や Web ストレージなどの技術を使用できます。こうした技術の詳細については、「[Excel-Add-in-JavaScript-PersistCustomSettings](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="02c83-p105">**Note:** Task pane add-ins for Project 2013 don't support the **Settings** API for storing add-in state or settings. However, for add-ins running in Project (as well as other Office host applications) you can use techniques such as browser cookies or web storage. For more information on these techniques, see the [Excel-Add-in-JavaScript-PersistCustomSettings](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings).</span></span> |
|[<span data-ttu-id="02c83-131">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="02c83-131">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings)|<span data-ttu-id="02c83-132">Outlook</span><span class="sxs-lookup"><span data-stu-id="02c83-132">Outlook</span></span>|<span data-ttu-id="02c83-133">アドインがインストールされている、ユーザーの Exchange サーバー メールボックス。これらの設定はユーザーのサーバー メールボックスに保存されるので、ユーザーと共に "ローミング" でき、そのユーザーのメールボックスにアクセスしている、サポートされているクライアント ホスト アプリケーションまたはブラウザーのコンテキストでアドインが実行されている場合、そのアドインでこれらの設定を利用できます。</span><span class="sxs-lookup"><span data-stu-id="02c83-133">The user's Exchange server mailbox where the add-in is installed.Because these settings are stored in the user's server mailbox, they can "roam" with the user and are available to the add-in when it is running in the context of any supported client host application or browser accessing that user's mailbox.</span></span><br/><br/> <span data-ttu-id="02c83-134">Outlook アドインのローミング設定は、その設定を作成したアドインのみが利用でき、また、アドインがインストールされているメールボックスからのみ利用できます。</span><span class="sxs-lookup"><span data-stu-id="02c83-134">Outlook add-in roaming settings are available only to the add-in that created them, and only from the mailbox where the add-in is installed.</span></span>|<span data-ttu-id="02c83-135">Outlook</span><span class="sxs-lookup"><span data-stu-id="02c83-135">Outlook</span></span>|
|[<span data-ttu-id="02c83-136">CustomProperties</span><span class="sxs-lookup"><span data-stu-id="02c83-136">CustomProperties</span></span>](/javascript/api/outlook/office.customproperties)|<span data-ttu-id="02c83-137">Outlook</span><span class="sxs-lookup"><span data-stu-id="02c83-137">Outlook</span></span>|<span data-ttu-id="02c83-p106">アドインが連携するメッセージ、予定、または会議出席依頼アイテム。 Outlook アドイン アイテムのカスタム プロパティは、そのプロパティを作成したアドインのみが利用でき、また、プロパティが保存されているアイテムからのみ利用できます。</span><span class="sxs-lookup"><span data-stu-id="02c83-p106">The message, appointment, or meeting request item the add-in is working with. Outlook add-in item custom properties are available only to the add-in that created them, and only from the item where they are saved.</span></span>|<span data-ttu-id="02c83-140">Outlook</span><span class="sxs-lookup"><span data-stu-id="02c83-140">Outlook</span></span>|
|[<span data-ttu-id="02c83-141">CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="02c83-141">CustomXmlParts</span></span>](/javascript/api/office/office.customxmlparts)|<span data-ttu-id="02c83-142">作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="02c83-142">task pane</span></span>|<span data-ttu-id="02c83-p107">アドインが連携しているドキュメント、スプレッドシート、またはプレゼンテーション。作業ウィンドウのアドイン設定は、その設定が保存されているドキュメントから、その設定を作成したアドインで使用できます。</span><span class="sxs-lookup"><span data-stu-id="02c83-p107">The document, spreadsheet, or presentation the add-in is working with. Task pane add-in settings are available to the add-in that created them from the document where they are saved.</span></span><br/><br/><span data-ttu-id="02c83-p108">**重要:** カスタム XML 部分には、パスワードなどの個人情報 (PII) を保存しないでください。保存されたデータはユーザーに対して表示されませんが、ドキュメントの一部として保存されるため、ドキュメントのファイル形式を直接読み取ることでアクセスできます。アドインによる PII の使用と、アドインが必要とするすべての PII の保存は、開発するアドインをユーザーのセキュリティが保護されるリソースとしてホストするサーバーのみで行うよう制限する必要があります。</span><span class="sxs-lookup"><span data-stu-id="02c83-p108">**Important:** Don't store passwords and other sensitive personally identifiable information (PII) in a custom XML part. The data saved isn't visible to end users, but it is stored as part of the document, which is accessible by reading the document's file format directly. You should limit your add-in's use of PII and store any PII required by your add-in only on the server hosting your add-in as a user-secured resource.</span></span>|<span data-ttu-id="02c83-148">Word (Office JavaScript 共通 API を使用)、Excel (ホスト固有の Excel JavaScript API を使用)</span><span class="sxs-lookup"><span data-stu-id="02c83-148">Word (using the Office JavaScript Common API) Excel (using the host-specific Excel JavaScript API</span></span>|

## <a name="settings-data-is-managed-in-memory-at-runtime"></a><span data-ttu-id="02c83-149">実行時のメモリ内での設定データの管理</span><span class="sxs-lookup"><span data-stu-id="02c83-149">Settings data is managed in memory at runtime</span></span>

> [!NOTE]
> <span data-ttu-id="02c83-p109">この後の 2 つのセクションでは、Office 共通 JavaScript API のコンテキストでの設定について説明します。 ホスト固有の Excel JavaScript API でも、カスタム設定にアクセスできます。 Excel の API とプログラミング パターンには、わずかな違いがあります。 詳細については、[Excel の SettingCollection](/javascript/api/excel/excel.settingcollection) を参照してください。</span><span class="sxs-lookup"><span data-stu-id="02c83-p109">The following two sections discuss settings in the context of the Office Common JavaScript API. The host-specific Excel JavaScript API also provides access to the custom settings. The Excel APIs and programming patterns are somewhat different. For more information, see [Excel SettingCollection](/javascript/api/excel/excel.settingcollection).</span></span>

<span data-ttu-id="02c83-154">内部的に、、、または`Settings` `CustomProperties` `RoamingSettings`オブジェクトでアクセスされるプロパティバッグ内のデータは、名前と値のペアを含むシリアル化された JavaScript object Notation (JSON) オブジェクトとして格納されます。</span><span class="sxs-lookup"><span data-stu-id="02c83-154">Internally, the data in the property bag accessed with the `Settings`, `CustomProperties`, or `RoamingSettings` objects is stored as a serialized JavaScript Object Notation (JSON) object that contains name/value pairs.</span></span> <span data-ttu-id="02c83-155">各`string`値の名前 (キー) は、である必要があります。また、格納さ`string`れ`number`た`date`値は`object`、**関数**ではなく、JavaScript、、、またはです。</span><span class="sxs-lookup"><span data-stu-id="02c83-155">The name (key) for each value must be a `string`, and the stored value can be a JavaScript `string`, `number`, `date`, or `object`, but not a **function**.</span></span>

<span data-ttu-id="02c83-156">この例はプロパティ バッグの構造を示し、3 つの定義された **string** 値 (`firstName`、`location`、`defaultView` という名前) が含まれます。</span><span class="sxs-lookup"><span data-stu-id="02c83-156">This example of the property bag structure contains three defined **string** values named `firstName`,  `location`, and  `defaultView`.</span></span>

```json
{
    "firstName":"Erik",
    "location":"98052",
    "defaultView":"basic"
}
```

<span data-ttu-id="02c83-157">設定プロパティ バッグは、前のアドイン セッション中に保存された後、アドインが初期化されるとき、またはその後はいつでも、アドインの現行セッション中は読み込むことができます。</span><span class="sxs-lookup"><span data-stu-id="02c83-157">After the settings property bag is saved during the previous add-in session, it can be loaded when the add-in is initialized or at any point after that during the add-in's current session.</span></span> <span data-ttu-id="02c83-158">セッション中は、作成する設定の種類 (**settings**、 **CustomProperties**、 `get`また`set`は**RoamingSettings**) に対応したオブジェクトの、、および`remove`メソッドを使用して、すべての設定がメモリ内で管理されます。</span><span class="sxs-lookup"><span data-stu-id="02c83-158">During the session, the settings are managed in entirely in memory using the `get`, `set`, and `remove` methods of the object that corresponds to the kind of settings you are creating (**Settings**, **CustomProperties**, or **RoamingSettings**).</span></span>


> [!IMPORTANT]
> <span data-ttu-id="02c83-159">アドインの現在のセッション中に行った追加、更新、または削除を保存場所に保持するには、その種類`saveAsync`の設定を操作するために使用される対応するオブジェクトのメソッドを呼び出す必要があります。</span><span class="sxs-lookup"><span data-stu-id="02c83-159">To persist any additions, updates, or deletions made during the add-in's current session to the storage location, you must call the `saveAsync` method of the corresponding object used to work with that kind of settings.</span></span> <span data-ttu-id="02c83-160">`set`、、および`remove`メソッドは`get`、設定プロパティバッグのメモリ内コピーに対してのみ動作します。</span><span class="sxs-lookup"><span data-stu-id="02c83-160">The `get`, `set`, and `remove` methods operate only on the in-memory copy of the settings property bag.</span></span> <span data-ttu-id="02c83-161">アドインを呼び出し`saveAsync`ずに閉じた場合、そのセッション中に設定に加えられた変更はすべて失われます。</span><span class="sxs-lookup"><span data-stu-id="02c83-161">If your add-in is closed without calling `saveAsync`, any changes made to settings during that session will be lost.</span></span>


## <a name="how-to-save-add-in-state-and-settings-per-document-for-content-and-task-pane-add-ins"></a><span data-ttu-id="02c83-162">コンテンツ アドインおよび作業ウィンドウ アドインで、ドキュメントごとにアドインの状態と設定を保存する方法</span><span class="sxs-lookup"><span data-stu-id="02c83-162">How to save add-in state and settings per document for content and task pane add-ins</span></span>


<span data-ttu-id="02c83-163">Word、Excel、または PowerPoint 用のコンテンツ アドインまたは作業ウィンドウ アドインの状態またはカスタム設定を保持するには、 [Settings](/javascript/api/office/office.settings) オブジェクトとそのメソッドを使用します。</span><span class="sxs-lookup"><span data-stu-id="02c83-163">To persist state or custom settings of a content or task pane add-in for Word, Excel, or PowerPoint, you use the [Settings](/javascript/api/office/office.settings) object and its methods.</span></span> <span data-ttu-id="02c83-164">`Settings`オブジェクトのメソッドを使用して作成されたプロパティバッグは、そのオブジェクトを作成したコンテンツまたは作業ウィンドウアドインのインスタンスのみが利用でき、保存されているドキュメントからのみ使用できます。</span><span class="sxs-lookup"><span data-stu-id="02c83-164">The property bag created with the methods of the `Settings` object are available only to the instance of the content or task pane add-in that created it, and only from the document in which it is saved.</span></span>

<span data-ttu-id="02c83-165">`Settings`オブジェクトは[Document](/javascript/api/office/office.document)オブジェクトの一部として自動的に読み込まれ、作業ウィンドウアドインまたはコンテンツアドインがアクティブ化されたときに使用できます。</span><span class="sxs-lookup"><span data-stu-id="02c83-165">The `Settings` object is automatically loaded as part of the [Document](/javascript/api/office/office.document) object, and is available when the task pane or content add-in is activated.</span></span> <span data-ttu-id="02c83-166">`Document`オブジェクトをインスタンス化した後、オブジェクトの`Settings` `Document` [settings](/javascript/api/office/office.document#settings)プロパティを使用して、そのオブジェクトにアクセスできます。</span><span class="sxs-lookup"><span data-stu-id="02c83-166">After the `Document` object is instantiated, you can access the `Settings` object with the [settings](/javascript/api/office/office.document#settings) property of the `Document` object.</span></span> <span data-ttu-id="02c83-167">セッションの有効期間中は`Settings.get`、 `Settings.set`プロパティバッグのメモリ内コピーにある`Settings.remove`永続化設定とアドイン状態を読み取り、書き込み、または削除するために、、、およびメソッドを使用するだけで済みます。</span><span class="sxs-lookup"><span data-stu-id="02c83-167">During the lifetime of the session, you can just use the `Settings.get`, `Settings.set`, and `Settings.remove` methods to read, write, or remove persisted settings and add-in state from the in-memory copy of the property bag.</span></span>

<span data-ttu-id="02c83-168">set メソッドと remove メソッドは設定プロパティ バッグのメモリ内コピーに対してのみ動作するので、アドインが関連付けられているドキュメントに新しい設定を保存、または変更された設定を保存し直すには [Settings.saveAsync](/javascript/api/office/office.settings#saveasync-options--callback-) メソッドを呼び出す必要があります。</span><span class="sxs-lookup"><span data-stu-id="02c83-168">Because the set and remove methods operate against only the in-memory copy of the settings property bag, to save new or changed settings back to the document the add-in is associated with you must call the [Settings.saveAsync](/javascript/api/office/office.settings#saveasync-options--callback-) method.</span></span>


### <a name="creating-or-updating-a-setting-value"></a><span data-ttu-id="02c83-169">設定値の作成または更新</span><span class="sxs-lookup"><span data-stu-id="02c83-169">Creating or updating a setting value</span></span>

<span data-ttu-id="02c83-p115">次のコード例では、[Settings.set](/javascript/api/office/office.settings#set-name--value-) メソッドを使用して `'themeColor'` という名前の設定を作成し、値 `'green'` を指定する方法を説明します。set メソッドの最初のパラメーターは、設定するか作成する設定の _name_ (Id) であり、これは大文字と小文字が区別されます。2 番目のパラメーターは、設定の _value_ です。</span><span class="sxs-lookup"><span data-stu-id="02c83-p115">The following code example shows how to use the [Settings.set](/javascript/api/office/office.settings#set-name--value-) method to create a setting called `'themeColor'` with a value `'green'`. The first parameter of the set method is the case-sensitive  _name_ (Id) of the setting to set or create. The second parameter is the _value_ of the setting.</span></span>


```js
Office.context.document.settings.set('themeColor', 'green');
```

 <span data-ttu-id="02c83-173">指定した名前を持つ設定は、それがまだ存在していない場合には作成され、すでに存在している場合はその値が更新されます。</span><span class="sxs-lookup"><span data-stu-id="02c83-173">The setting with the specified name is created if it doesn't already exist, or its value is updated if it does exist.</span></span> <span data-ttu-id="02c83-174">`Settings.saveAsync`メソッドを使用して、新しい設定または更新された設定をドキュメントに保持します。</span><span class="sxs-lookup"><span data-stu-id="02c83-174">Use the `Settings.saveAsync` method to persist the new or updated settings to the document.</span></span>


### <a name="getting-the-value-of-a-setting"></a><span data-ttu-id="02c83-175">設定値の取得</span><span class="sxs-lookup"><span data-stu-id="02c83-175">Getting the value of a setting</span></span>

<span data-ttu-id="02c83-176">次の例では、 [Settings.get](/javascript/api/office/office.settings#get-name-) メソッドを使用して "themeColor" という名前の設定値を取得する方法を示します。</span><span class="sxs-lookup"><span data-stu-id="02c83-176">The following example shows how use the [Settings.get](/javascript/api/office/office.settings#get-name-) method to get the value of a setting called "themeColor".</span></span> <span data-ttu-id="02c83-177">この`get`メソッドの唯一のパラメーターは、大文字と小文字が区別される設定の_名前_です。</span><span class="sxs-lookup"><span data-stu-id="02c83-177">The only parameter of the `get` method is the case-sensitive _name_ of the setting.</span></span>


```js
write('Current value for mySetting: ' + Office.context.document.settings.get('themeColor'));

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

 <span data-ttu-id="02c83-178">メソッド`get`は、渡された設定_名_に対して以前に保存された値を返します。</span><span class="sxs-lookup"><span data-stu-id="02c83-178">The `get` method returns the value that was previously saved for the setting _name_ that was passed in.</span></span> <span data-ttu-id="02c83-179">設定が存在しない場合、メソッドは **null** を返します。</span><span class="sxs-lookup"><span data-stu-id="02c83-179">If the setting doesn't exist, the method returns **null**.</span></span>


### <a name="removing-a-setting"></a><span data-ttu-id="02c83-180">設定の削除</span><span class="sxs-lookup"><span data-stu-id="02c83-180">Removing a setting</span></span>

<span data-ttu-id="02c83-181">次の例では、 [Settings.remove](/javascript/api/office/office.settings#remove-name-) メソッドを使用して、"themeColor" という名前の設定を削除する方法を示します。</span><span class="sxs-lookup"><span data-stu-id="02c83-181">The following example shows how to use the [Settings.remove](/javascript/api/office/office.settings#remove-name-) method to remove a setting with the name "themeColor".</span></span> <span data-ttu-id="02c83-182">この`remove`メソッドの唯一のパラメーターは、大文字と小文字が区別される設定の_名前_です。</span><span class="sxs-lookup"><span data-stu-id="02c83-182">The only parameter of the `remove` method is the case-sensitive _name_ of the setting.</span></span>


```js
Office.context.document.settings.remove('themeColor');
```

<span data-ttu-id="02c83-183">該当する設定が存在しない場合は何も起きません。</span><span class="sxs-lookup"><span data-stu-id="02c83-183">Nothing will happen if the setting does not exist.</span></span> <span data-ttu-id="02c83-184">`Settings.saveAsync`メソッドを使用して、ドキュメントから設定の削除を保持します。</span><span class="sxs-lookup"><span data-stu-id="02c83-184">Use the `Settings.saveAsync` method to persist removal of the setting from the document.</span></span>


### <a name="saving-your-settings"></a><span data-ttu-id="02c83-185">設定の保存</span><span class="sxs-lookup"><span data-stu-id="02c83-185">Saving your settings</span></span>

<span data-ttu-id="02c83-186">現在のセッション中に、アドインがメモリ内の設定プロパティ バッグに対して行った追加、変更、または削除を保存するには、 [Settings.saveAsync](/javascript/api/office/office.settings#saveasync-options--callback-) メソッドを呼び出してそれらの設定をドキュメントに保存する必要があります。</span><span class="sxs-lookup"><span data-stu-id="02c83-186">To save any additions, changes, or deletions your add-in made to the in-memory copy of the settings property bag during the current session, you must call the [Settings.saveAsync](/javascript/api/office/office.settings#saveasync-options--callback-) method to store them in the document.</span></span> <span data-ttu-id="02c83-187">`saveAsync`メソッドの唯一のパラメーターは_callback_で、これは1つのパラメーターを持つコールバック関数です。</span><span class="sxs-lookup"><span data-stu-id="02c83-187">The only parameter of the `saveAsync` method is _callback_, which is a callback function with a single parameter.</span></span> 


```js
Office.context.document.settings.saveAsync(function (asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        write('Settings save failed. Error: ' + asyncResult.error.message);
    } else {
        write('Settings saved.');
    }
});
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

<span data-ttu-id="02c83-188">`saveAsync`メソッドに_コールバック_パラメーターとして渡される匿名関数は、操作が完了したときに実行されます。</span><span class="sxs-lookup"><span data-stu-id="02c83-188">The anonymous function passed into the `saveAsync` method as the _callback_ parameter is executed when the operation is completed.</span></span> <span data-ttu-id="02c83-189">コールバックの_asyncResult_パラメーターは、操作の状態`AsyncResult`を含むオブジェクトへのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="02c83-189">The _asyncResult_ parameter of the callback provides access to an `AsyncResult` object that contains the status of the operation.</span></span> <span data-ttu-id="02c83-190">この例では、関数は`AsyncResult.status`プロパティをチェックして、保存操作が成功したか失敗したかを確認し、アドインのページに結果を表示します。</span><span class="sxs-lookup"><span data-stu-id="02c83-190">In the example, the function checks the `AsyncResult.status` property to see if the save operation succeeded or failed, and then displays the result in the add-in's page.</span></span>

## <a name="how-to-save-custom-xml-to-the-document"></a><span data-ttu-id="02c83-191">ドキュメントにカスタム XML を保存する方法</span><span class="sxs-lookup"><span data-stu-id="02c83-191">How to save custom XML to the document</span></span>

> [!NOTE]
> <span data-ttu-id="02c83-p123">このセクションでは、Word でサポートされている Office 共通 JavaScript API のコンテキストでのカスタム XML 部分について説明します。 ホスト固有の Excel JavaScript API でも、カスタム XML 部分にアクセスできます。 Excel の API とプログラミング パターンには、わずかな違いがあります。 詳細については、[Excel の CustomXmlPart](/javascript/api/excel/excel.customxmlpart) を参照してください。</span><span class="sxs-lookup"><span data-stu-id="02c83-p123">This section discusses custom XML parts in the context of the Office Common JavaScript API which is supported in Word. The host-specific Excel JavaScript API also provides access to the custom XML parts. The Excel APIs and programming patterns are somewhat different. For more information, see [Excel CustomXmlPart](/javascript/api/excel/excel.customxmlpart).</span></span>

<span data-ttu-id="02c83-196">ドキュメントの Settings のサイズ制限を超過する情報や構造化された特徴を持つ情報を保存する必要がある場合には、追加のストレージ オプションがあります。</span><span class="sxs-lookup"><span data-stu-id="02c83-196">There is an addtional storage option when you need to store information that exceeds the size limits of the document Settings or which has a structured character.</span></span> <span data-ttu-id="02c83-197">Word および Excel の作業ウィンドウ アドインには、カスタムの XML マークアップを保持できます (Excel については、このセクションの冒頭にあるノートを参照してください)。</span><span class="sxs-lookup"><span data-stu-id="02c83-197">You can persist custom XML markup in a task pane add-in for Word (and for Excel, but see the note at the top of this section).</span></span> <span data-ttu-id="02c83-198">Word の場合は、[CustomXmlPart](/javascript/api/office/office.customxmlpart) とそのメソッドを使用します (繰り返しになりますが、Excel の場合は上記のノートを参照してください)。</span><span class="sxs-lookup"><span data-stu-id="02c83-198">In Word, you use the [CustomXmlPart](/javascript/api/office/office.customxmlpart) object and its methods (again, see the note above for Excel).</span></span> <span data-ttu-id="02c83-199">次のコードでは、カスタム XML パーツを作成して、その ID とコンテンツをページの div に表示します。</span><span class="sxs-lookup"><span data-stu-id="02c83-199">The following code creates a custom XML part and displays its ID and then its content in divs on the page.</span></span> <span data-ttu-id="02c83-200">XML 文字列には `xmlns` 属性が必ず存在する点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="02c83-200">Note that there must be an `xmlns` attribute in the XML string.</span></span>

```js
function createCustomXmlPart() {
    const xmlString = "<Reviewers xmlns='http://schemas.contoso.com/review/1.0'><Reviewer>Juan</Reviewer><Reviewer>Hong</Reviewer><Reviewer>Sally</Reviewer></Reviewers>";
    Office.context.document.customXmlParts.addAsync(xmlString,
        (asyncResult) => {
            $("#xml-id").text("Your new XML part's ID: " + asyncResult.value.id);
            asyncResult.value.getXmlAsync(
                (asyncResult) => {
                    $("#xml-blob").text(asyncResult.value);
                }
            );
        }
    );
}
```

<span data-ttu-id="02c83-p125">カスタム XML 部分を取得するには、[getByIdAsync](/javascript/api/office/office.customxmlparts#getbyidasync-id--options--callback-) メソッドを使用しますが、ID は XML 部分の作成時に生成された GUID になるため、コードの作成時に ID の内容を知ることはできません。 そのため、XML 部分を作成したら、その XML 部分の ID を設定としてすぐに保存して、覚えやすいキーを割り当てることがベスト プラクティスになります。 次のメソッドは、この方法を示してます  (ただし、カスタム設定の操作に関する詳細とベスト プラクティスについては、この記事の前半のセクションを参照してください)。</span><span class="sxs-lookup"><span data-stu-id="02c83-p125">To retrieve a custom XML part, you use the [getByIdAsync](/javascript/api/office/office.customxmlparts#getbyidasync-id--options--callback-) method, but the ID is a GUID that is generated when the XML part is created, so you can't know when coding what the ID is. For that reason, it is a good practice when creating an XML part to immediately store the ID of the XML part as a setting and give it a memorable key. The following method shows how to do this. (But see earlier sections of this article for details and best practices when working with custom settings).</span></span>

 ```js
function createCustomXmlPartAndStoreId() {
    const xmlString = "<Reviewers xmlns='http://schemas.contoso.com/review/1.0'><Reviewer>Juan</Reviewer><Reviewer>Hong</Reviewer><Reviewer>Sally</Reviewer></Reviewers>";
    Office.context.document.customXmlParts.addAsync(xmlString,
        (asyncResult) => {
            Office.context.document.settings.set('ReviewersID', asyncResult.id);
            Office.context.document.settings.saveAsync();
        }
    );
}
```

<span data-ttu-id="02c83-205">次のコードは、最初に設定から ID を取得することで、XML 部分を取得する方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="02c83-205">The following code shows how to retrieve the XML part by first getting its ID from a setting.</span></span>

 ```js
function getReviewers() {
    const reviewersXmlId = Office.context.document.settings.get('ReviewersID');
    Office.context.document.customXmlParts.getByIdAsync(reviewersXmlId,
        (asyncResult) => {
            asyncResult.value.getXmlAsync(
                (asyncResult) => {
                    $("#xml-blob").text(asyncResult.value);
                }
            );
        }
    );
}
```

## <a name="how-to-save-settings-in-an-outlook-add-in"></a><span data-ttu-id="02c83-206">Outlook アドインに設定を保存する方法</span><span class="sxs-lookup"><span data-stu-id="02c83-206">How to save settings in an Outlook add-in</span></span>

<span data-ttu-id="02c83-207">Outlook アドインに設定を保存する方法については、「 [outlook アドインの状態と設定を管理](../outlook/manage-state-and-settings-outlook.md)する」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="02c83-207">For information about how to save settings in an Outlook add-in, see [Manage state and settings for an Outlook add-in](../outlook/manage-state-and-settings-outlook.md).</span></span>


## <a name="see-also"></a><span data-ttu-id="02c83-208">関連項目</span><span class="sxs-lookup"><span data-stu-id="02c83-208">See also</span></span>

- [<span data-ttu-id="02c83-209">Office JavaScript API について</span><span class="sxs-lookup"><span data-stu-id="02c83-209">Understanding the Office JavaScript API</span></span>](understanding-the-javascript-api-for-office.md)
- [<span data-ttu-id="02c83-210">Outlook アドイン</span><span class="sxs-lookup"><span data-stu-id="02c83-210">Outlook add-ins</span></span>](../outlook/outlook-add-ins-overview.md)
- [<span data-ttu-id="02c83-211">Outlook アドインの状態と設定を管理する</span><span class="sxs-lookup"><span data-stu-id="02c83-211">Manage state and settings for an Outlook add-in</span></span>](../outlook/manage-state-and-settings-outlook.md)
- [<span data-ttu-id="02c83-212">Excel-Add-in-JavaScript-PersistCustomSettings</span><span class="sxs-lookup"><span data-stu-id="02c83-212">Excel-Add-in-JavaScript-PersistCustomSettings</span></span>](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings)
