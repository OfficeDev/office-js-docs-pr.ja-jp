---
title: アドインの状態および設定を保持する
description: ''
ms.date: 03/19/2019
localization_priority: Priority
ms.openlocfilehash: 6092a93751825561f83cfea1671fe59e273f6142
ms.sourcegitcommit: 9c5a836d4464e49846c9795bf44cfe23e9fc8fbe
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/10/2019
ms.locfileid: "35617019"
---
# <a name="persisting-add-in-state-and-settings"></a><span data-ttu-id="63333-102">アドインの状態および設定を保持する</span><span class="sxs-lookup"><span data-stu-id="63333-102">Persisting add-in state and settings</span></span>

<span data-ttu-id="63333-p101">Office アドインは、基本的にブラウザー コントロールのステートレス環境で動作する Web アプリケーションです。したがって、アドインでは、そのアドインを使用するセッション間で特定の操作または機能を継続して維持するためのデータを保持することが必要な場合があります。たとえば、アドインには、ユーザーの優先ビューや既定の場所など、アドインで保存しておき、アドインが次回初期化されたときにリロードする必要があるカスタム設定または他の値が含まれる場合があります。その場合は、次のようにします。</span><span class="sxs-lookup"><span data-stu-id="63333-p101">Office Add-ins are essentially web applications running in the stateless environment of a browser control. As a result, your add-in may need to persist data to maintain the continuity of certain operations or features across sessions of using your add-in. For example, your add-in may have custom settings or other values that it needs to save and reload the next time it's initialized, such as a user's preferred view or default location. To do that, you can:</span></span>

- <span data-ttu-id="63333-107">次のどちらかの形式でデータを保存する JavaScript API for Office のメンバー使用します。</span><span class="sxs-lookup"><span data-stu-id="63333-107">Use members of the JavaScript API for Office that store data as either:</span></span>
    -  <span data-ttu-id="63333-108">アドインの種類応じた場所に保存されるプロパティ バッグ内の名前と値の組。</span><span class="sxs-lookup"><span data-stu-id="63333-108">Name/value pairs in a property bag stored in a location that depends on add-in type.</span></span>
    -  <span data-ttu-id="63333-109">ドキュメント内に保存されるカスタム XML。</span><span class="sxs-lookup"><span data-stu-id="63333-109">Custom XML stored in the document.</span></span>

- <span data-ttu-id="63333-110">基になるブラウザー コントロールによって提供される技術である、ブラウザーの Cookie、または HTML5 Web ストレージ ([localStorage](https://developer.mozilla.org/docs/Web/API/Window/localStorage) または [sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage)) を使用します。</span><span class="sxs-lookup"><span data-stu-id="63333-110">Use techniques provided by the underlying browser control: browser cookies, or HTML5 web storage ([localStorage](https://developer.mozilla.org/docs/Web/API/Window/localStorage) or [sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage)).</span></span>

<span data-ttu-id="63333-p102">この記事では、アドインの状態を保持する JavaScript API for Office の使い方について説明します。ブラウザーの Cookie および Web ストレージの使用例については、「 [Excel-Add-in-JavaScript-PersistCustomSettings](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="63333-p102">This article focuses on how to use the JavaScript API for Office to persist add-in state. For examples of using browser cookies and web storage, see the [Excel-Add-in-JavaScript-PersistCustomSettings](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings).</span></span>

## <a name="persisting-add-in-state-and-settings-with-the-javascript-api-for-office"></a><span data-ttu-id="63333-113">JavaScript API for Office を使用してアドインの状態および設定を保持する</span><span class="sxs-lookup"><span data-stu-id="63333-113">Persisting add-in state and settings with the JavaScript API for Office</span></span>

<span data-ttu-id="63333-p103">JavaScript API for Office には、次の表に示すように、セッション間でアドインの状態を保存するために [Settings](/javascript/api/office/office.settings) オブジェクト、 [RoamingSettings](/javascript/api/outlook/office.roamingsettings) オブジェクト、および [CustomProperties](/javascript/api/outlook/office.customproperties) オブジェクトが用意されています。すべてのケースで、保存された設定値は、それを作成したアドインの [Id](/office/dev/add-ins/reference/manifest/id) にのみ関連付けられます。</span><span class="sxs-lookup"><span data-stu-id="63333-p103">The JavaScript API for Office provides the [Settings](/javascript/api/office/office.settings), [RoamingSettings](/javascript/api/outlook/office.roamingsettings), and [CustomProperties](/javascript/api/outlook/office.customproperties) objects for saving add-in state across sessions as described in the following table. In all cases, the saved settings values are associated with the [Id](/office/dev/add-ins/reference/manifest/id) of the add-in that created them.</span></span>

|<span data-ttu-id="63333-116">**オブジェクト**</span><span class="sxs-lookup"><span data-stu-id="63333-116">**Object**</span></span>|<span data-ttu-id="63333-117">**アドインの種類のサポート**</span><span class="sxs-lookup"><span data-stu-id="63333-117">**Add-in type support**</span></span>|<span data-ttu-id="63333-118">**ストレージの場所**</span><span class="sxs-lookup"><span data-stu-id="63333-118">**Storage location**</span></span>|<span data-ttu-id="63333-119">**サポートされる Office のホスト**</span><span class="sxs-lookup"><span data-stu-id="63333-119">**Office host support**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="63333-120">Settings</span><span class="sxs-lookup"><span data-stu-id="63333-120">Settings</span></span>](/javascript/api/office/office.settings)|<span data-ttu-id="63333-121">コンテンツおよび作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="63333-121">content and task pane</span></span>|<span data-ttu-id="63333-122">アドインが連携しているドキュメント、スプレッドシート、またはプレゼンテーション。コンテンツおよび作業ウィンドウのアドイン設定は、その設定が保存されているドキュメントから、その設定を作成したアドインで使用できます。</span><span class="sxs-lookup"><span data-stu-id="63333-122">The document, spreadsheet, or presentation the add-in is working with.Content and task pane add-in settings are available to the add-in that created them from the document where they are saved.</span></span><br/><br/><span data-ttu-id="63333-p104">**重要:\*\*\*\*Settings** オブジェクトを使用して、パスワードおよびその他の機密の個人を特定できる情報 (PII) を保存しないでください。保存されたデータはユーザーに対して表示されませんが、ドキュメントの一部として保存されているため、ドキュメントのファイル形式を直接読み取ることでアクセスできます。アドインによる PII の使用と、アドインが必要とするすべての PII の保存は、開発するアドインをユーザーのセキュリティが保護されるリソースとしてホストするサーバーのみで行うよう制限する必要があります。</span><span class="sxs-lookup"><span data-stu-id="63333-p104">**Important:** Don't store passwords and other sensitive personally identifiable information (PII) with the **Settings** object. The data saved isn't visible to end users, but it is stored as part of the document, which is accessible by reading the document's file format directly. You should limit your add-in's use of PII and store any PII required by your add-in only on the server hosting your add-in as a user-secured resource.</span></span>|<span data-ttu-id="63333-126">Word、Excel、または PowerPoint</span><span class="sxs-lookup"><span data-stu-id="63333-126">Word, Excel, or PowerPoint</span></span><br/><br/> <span data-ttu-id="63333-p105">**メモ:** Project 2013 の作業ウィンドウ アドインでは、アドインの状態または設定を保存するための **Settings** API をサポートしていません。ただし、Project (および他の Office ホスト アプリケーション) で動作するアドインの場合は、ブラウザーの Cookie や Web ストレージなどの技術を使用できます。こうした技術の詳細については、「[Excel-Add-in-JavaScript-PersistCustomSettings](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="63333-p105">**Note:** Task pane add-ins for Project 2013 don't support the **Settings** API for storing add-in state or settings. However, for add-ins running in Project (as well as other Office host applications) you can use techniques such as browser cookies or web storage. For more information on these techniques, see the [Excel-Add-in-JavaScript-PersistCustomSettings](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings).</span></span> |
|[<span data-ttu-id="63333-130">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="63333-130">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings)|<span data-ttu-id="63333-131">Outlook</span><span class="sxs-lookup"><span data-stu-id="63333-131">Outlook</span></span>|<span data-ttu-id="63333-132">アドインがインストールされている、ユーザーの Exchange サーバー メールボックス。これらの設定はユーザーのサーバー メールボックスに保存されるので、ユーザーと共に "ローミング" でき、そのユーザーのメールボックスにアクセスしている、サポートされているクライアント ホスト アプリケーションまたはブラウザーのコンテキストでアドインが実行されている場合、そのアドインでこれらの設定を利用できます。</span><span class="sxs-lookup"><span data-stu-id="63333-132">The user's Exchange server mailbox where the add-in is installed.Because these settings are stored in the user's server mailbox, they can "roam" with the user and are available to the add-in when it is running in the context of any supported client host application or browser accessing that user's mailbox.</span></span><br/><br/> <span data-ttu-id="63333-133">Outlook アドインのローミング設定は、その設定を作成したアドインのみが利用でき、また、アドインがインストールされているメールボックスからのみ利用できます。</span><span class="sxs-lookup"><span data-stu-id="63333-133">Outlook add-in roaming settings are available only to the add-in that created them, and only from the mailbox where the add-in is installed.</span></span>|<span data-ttu-id="63333-134">Outlook</span><span class="sxs-lookup"><span data-stu-id="63333-134">Outlook</span></span>|
|[<span data-ttu-id="63333-135">CustomProperties</span><span class="sxs-lookup"><span data-stu-id="63333-135">CustomProperties</span></span>](/javascript/api/outlook/office.customproperties)|<span data-ttu-id="63333-136">Outlook</span><span class="sxs-lookup"><span data-stu-id="63333-136">Outlook</span></span>|<span data-ttu-id="63333-p106">アドインが連携するメッセージ、予定、または会議出席依頼アイテム。 Outlook アドイン アイテムのカスタム プロパティは、そのプロパティを作成したアドインのみが利用でき、また、プロパティが保存されているアイテムからのみ利用できます。</span><span class="sxs-lookup"><span data-stu-id="63333-p106">The message, appointment, or meeting request item the add-in is working with. Outlook add-in item custom properties are available only to the add-in that created them, and only from the item where they are saved.</span></span>|<span data-ttu-id="63333-139">Outlook</span><span class="sxs-lookup"><span data-stu-id="63333-139">Outlook</span></span>|
|[<span data-ttu-id="63333-140">CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="63333-140">CustomXmlParts</span></span>](/javascript/api/office/office.customxmlparts)|<span data-ttu-id="63333-141">作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="63333-141">task pane</span></span>|<span data-ttu-id="63333-p107">アドインが連携しているドキュメント、スプレッドシート、またはプレゼンテーション。作業ウィンドウのアドイン設定は、その設定が保存されているドキュメントから、その設定を作成したアドインで使用できます。</span><span class="sxs-lookup"><span data-stu-id="63333-p107">The document, spreadsheet, or presentation the add-in is working with. Task pane add-in settings are available to the add-in that created them from the document where they are saved.</span></span><br/><br/><span data-ttu-id="63333-p108">**重要:** カスタム XML 部分には、パスワードなどの個人情報 (PII) を保存しないでください。保存されたデータはユーザーに対して表示されませんが、ドキュメントの一部として保存されるため、ドキュメントのファイル形式を直接読み取ることでアクセスできます。アドインによる PII の使用と、アドインが必要とするすべての PII の保存は、開発するアドインをユーザーのセキュリティが保護されるリソースとしてホストするサーバーのみで行うよう制限する必要があります。</span><span class="sxs-lookup"><span data-stu-id="63333-p108">**Important:** Don't store passwords and other sensitive personally identifiable information (PII) in a custom XML part. The data saved isn't visible to end users, but it is stored as part of the document, which is accessible by reading the document's file format directly. You should limit your add-in's use of PII and store any PII required by your add-in only on the server hosting your add-in as a user-secured resource.</span></span>|<span data-ttu-id="63333-147">Word (Office JavaScript 共通 API を使用)、Excel (ホスト固有の Excel JavaScript API を使用)</span><span class="sxs-lookup"><span data-stu-id="63333-147">Word (using the Office JavaScript Common API) Excel (using the host-specific Excel JavaScript API</span></span>|

## <a name="settings-data-is-managed-in-memory-at-runtime"></a><span data-ttu-id="63333-148">実行時のメモリ内での設定データの管理</span><span class="sxs-lookup"><span data-stu-id="63333-148">Settings data is managed in memory at runtime</span></span>

> [!NOTE]
> <span data-ttu-id="63333-p109">この後の 2 つのセクションでは、Office 共通 JavaScript API のコンテキストでの設定について説明します。 ホスト固有の Excel JavaScript API でも、カスタム設定にアクセスできます。 Excel の API とプログラミング パターンには、わずかな違いがあります。 詳細については、[Excel の SettingCollection](/javascript/api/excel/excel.settingcollection) を参照してください。</span><span class="sxs-lookup"><span data-stu-id="63333-p109">The following two sections discuss settings in the context of the Office Common JavaScript API. The host-specific Excel JavaScript API also provides access to the custom settings. The Excel APIs and programming patterns are somewhat different. For more information, see [Excel SettingCollection](/javascript/api/excel/excel.settingcollection).</span></span>

<span data-ttu-id="63333-153">内部的には、**Settings** オブジェクト、**CustomProperties** オブジェクト、または **RoamingSettings** オブジェクトでアクセスされるプロパティ バッグ内のデータは、名前/値のペアを含むシリアル化された JavaScript Object Notation (JSON) オブジェクトとして格納されます。</span><span class="sxs-lookup"><span data-stu-id="63333-153">Internally, the data in the property bag accessed with the **Settings**, **CustomProperties**, or **RoamingSettings** objects is stored as a serialized JavaScript Object Notation (JSON) object that contains name/value pairs.</span></span> <span data-ttu-id="63333-154">各値の名前 (キー) は **string** である必要があり、格納された値は JavaScript の **string**、**number**、**date**、または **object** にすることが可能ですが、**function** にすることはできません。</span><span class="sxs-lookup"><span data-stu-id="63333-154">The name (key) for each value must be a **string**, and the stored value can be a JavaScript **string**, **number**, **date**, or **object**, but not a **function**.</span></span>

<span data-ttu-id="63333-155">この例はプロパティ バッグの構造を示し、3 つの定義された **string** 値 (`firstName`、`location`、`defaultView` という名前) が含まれます。</span><span class="sxs-lookup"><span data-stu-id="63333-155">This example of the property bag structure contains three defined **string** values named `firstName`,  `location`, and  `defaultView`.</span></span>

```json
{
    "firstName":"Erik",
    "location":"98052",
    "defaultView":"basic"
}
```

<span data-ttu-id="63333-156">設定プロパティ バッグは、前のアドイン セッション中に保存された後、アドインが初期化されるとき、またはその後はいつでも、アドインの現行セッション中は読み込むことができます。</span><span class="sxs-lookup"><span data-stu-id="63333-156">After the settings property bag is saved during the previous add-in session, it can be loaded when the add-in is initialized or at any point after that during the add-in's current session.</span></span> <span data-ttu-id="63333-157">セッションの間、作成している設定の種類に対応するオブジェクト (**Settings**、**CustomProperties**、**RoamingSettings**) の **get**、**set**、**remove** メソッドを使用し、メモリ内で設定全体が管理されます。</span><span class="sxs-lookup"><span data-stu-id="63333-157">During the session, the settings are managed in entirely in memory using the **get**, **set**, and **remove** methods of the object that corresponds to the kind settings you are creating ( **Settings**, **CustomProperties**, or **RoamingSettings**).</span></span> 


> [!IMPORTANT]
> <span data-ttu-id="63333-158">アドインの現行セッション中に行われた追加、更新、または削除を保存場所に保持するには、その種の設定の操作で使用される、対応するオブジェクトの **saveAsync** メソッドを呼び出す必要があります。</span><span class="sxs-lookup"><span data-stu-id="63333-158">To persist any additions, updates, or deletions made during the add-in's current session to the storage location, you must call the **saveAsync** method of the corresponding object used to work with that kind of settings.</span></span> <span data-ttu-id="63333-159">**get**、**set**、**remove** メソッドは、設定プロパティ バッグのメモリ内コピーでのみ動作します。</span><span class="sxs-lookup"><span data-stu-id="63333-159">The **get**, **set**, and **remove** methods operate only on the in-memory copy of the settings property bag.</span></span> <span data-ttu-id="63333-160">**saveAsync** の呼び出しなしにアドインが閉じられた場合、そのセッション中に設定に対して行われた変更は失われます。</span><span class="sxs-lookup"><span data-stu-id="63333-160">If your add-in is closed without calling **saveAsync**, any changes made to settings during that session will be lost.</span></span> 


## <a name="how-to-save-add-in-state-and-settings-per-document-for-content-and-task-pane-add-ins"></a><span data-ttu-id="63333-161">コンテンツ アドインおよび作業ウィンドウ アドインで、ドキュメントごとにアドインの状態と設定を保存する方法</span><span class="sxs-lookup"><span data-stu-id="63333-161">How to save add-in state and settings per document for content and task pane add-ins</span></span>


<span data-ttu-id="63333-p113">Word、Excel、または PowerPoint 用のコンテンツ アドインまたは作業ウィンドウ アドインの状態またはカスタム設定を保持するには、[Settings](/javascript/api/office/office.settings) オブジェクトとそのメソッドを使用します。**Settings** オブジェクトのメソッドを使用して作成されたプロパティ バッグは、それを作成したコンテンツ アドインまたは作業ウィンドウ アドインのインスタンスのみが利用でき、プロパティ バッグが保存されているドキュメント以外からは使用できません。</span><span class="sxs-lookup"><span data-stu-id="63333-p113">To persist state or custom settings of a content or task pane add-in for Word, Excel, or PowerPoint, you use the [Settings](/javascript/api/office/office.settings) object and its methods. The property bag created with the methods of the **Settings** object are available only to the instance of the content or task pane add-in that created it, and only from the document in which it is saved.</span></span>

<span data-ttu-id="63333-164">**Settings** オブジェクトは、[Document](/javascript/api/office/office.document) オブジェクトの一部として自動的に読み込まれ、作業ウィンドウまたはコンテンツ アドインがアクティブ化されると使用できるようになります。</span><span class="sxs-lookup"><span data-stu-id="63333-164">The **Settings** object is automatically loaded as part of the [Document](/javascript/api/office/office.document) object, and is available when the task pane or content add-in is activated.</span></span> <span data-ttu-id="63333-165">**Document** オブジェクトがインスタンス化された後は、**Document** オブジェクトの [settings](/javascript/api/office/office.document#settings) プロパティを使用して、**Settings** オブジェクトにアクセスできます。</span><span class="sxs-lookup"><span data-stu-id="63333-165">After the **Document** object is instantiated, you can access the **Settings** object with the [settings](/javascript/api/office/office.document#settings) property of the **Document** object.</span></span> <span data-ttu-id="63333-166">セッションの期間中は、**Settings.get**、**Settings.set**、**Settings.remove** メソッドを使用するだけで、永続的な設定およびアドインの状態の読み取り、書き込み、または削除をプロパティ バッグのメモリ内コピーで行うことができます。</span><span class="sxs-lookup"><span data-stu-id="63333-166">During the lifetime of the session, you can just use the **Settings.get**, **Settings.set**, and **Settings.remove** methods to read, write, or remove persisted settings and add-in state from the in-memory copy of the property bag.</span></span>

<span data-ttu-id="63333-167">set メソッドと remove メソッドは設定プロパティ バッグのメモリ内コピーに対してのみ動作するので、アドインが関連付けられているドキュメントに新しい設定を保存、または変更された設定を保存し直すには [Settings.saveAsync](/javascript/api/office/office.settings#saveasync-options--callback-) メソッドを呼び出す必要があります。</span><span class="sxs-lookup"><span data-stu-id="63333-167">Because the set and remove methods operate against only the in-memory copy of the settings property bag, to save new or changed settings back to the document the add-in is associated with you must call the [Settings.saveAsync](/javascript/api/office/office.settings#saveasync-options--callback-) method.</span></span>


### <a name="creating-or-updating-a-setting-value"></a><span data-ttu-id="63333-168">設定値の作成または更新</span><span class="sxs-lookup"><span data-stu-id="63333-168">Creating or updating a setting value</span></span>

<span data-ttu-id="63333-p115">次のコード例では、[Settings.set](/javascript/api/office/office.settings#set-name--value-) メソッドを使用して `'themeColor'` という名前の設定を作成し、値 `'green'` を指定する方法を説明します。set メソッドの最初のパラメーターは、設定するか作成する設定の _name_ (Id) であり、これは大文字と小文字が区別されます。2 番目のパラメーターは、設定の _value_ です。</span><span class="sxs-lookup"><span data-stu-id="63333-p115">The following code example shows how to use the [Settings.set](/javascript/api/office/office.settings#set-name--value-) method to create a setting called `'themeColor'` with a value `'green'`. The first parameter of the set method is the case-sensitive  _name_ (Id) of the setting to set or create. The second parameter is the _value_ of the setting.</span></span>


```js
Office.context.document.settings.set('themeColor', 'green');
```

 <span data-ttu-id="63333-p116">指定した名前を持つ設定は、それがまだ存在していない場合には作成され、すでに存在している場合はその値が更新されます。**Settings.saveAsync** メソッドを使用すると、新しい設定または更新された設定をドキュメントに保持できます。</span><span class="sxs-lookup"><span data-stu-id="63333-p116">The setting with the specified name is created if it doesn't already exist, or its value is updated if it does exist. Use the **Settings.saveAsync** method to persist the new or updated settings to the document.</span></span>


### <a name="getting-the-value-of-a-setting"></a><span data-ttu-id="63333-174">設定値の取得</span><span class="sxs-lookup"><span data-stu-id="63333-174">Getting the value of a setting</span></span>

<span data-ttu-id="63333-p117">次の例では、[Settings.get](/javascript/api/office/office.settings#get-name-) メソッドを使用して "themeColor" という名前の設定値を取得する方法を示します。**get** メソッドの唯一のパラメーターは、設定の _name_ であり、これは大文字と小文字が区別されます。</span><span class="sxs-lookup"><span data-stu-id="63333-p117">The following example shows how use the [Settings.get](/javascript/api/office/office.settings#get-name-) method to get the value of a setting called "themeColor". The only parameter of the **get** method is the case-sensitive _name_ of the setting.</span></span>


```js
write('Current value for mySetting: ' + Office.context.document.settings.get('themeColor'));

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

 <span data-ttu-id="63333-p118">**get** メソッドでは、指定した _name_ という設定に対して以前に保存した値を返します。設定が存在しない場合、メソッドは **null** を返します。</span><span class="sxs-lookup"><span data-stu-id="63333-p118">The **get** method returns the value that was previously saved for the setting _name_ that was passed in. If the setting doesn't exist, the method returns **null**.</span></span>


### <a name="removing-a-setting"></a><span data-ttu-id="63333-179">設定の削除</span><span class="sxs-lookup"><span data-stu-id="63333-179">Removing a setting</span></span>

<span data-ttu-id="63333-p119">次の例では、[Settings.remove](/javascript/api/office/office.settings#remove-name-) メソッドを使用して、"themeColor" という名前の設定を削除する方法を示します。**remove** メソッドの唯一のパラメーターは設定の _name_ であり、これは大文字と小文字が区別されます。</span><span class="sxs-lookup"><span data-stu-id="63333-p119">The following example shows how to use the [Settings.remove](/javascript/api/office/office.settings#remove-name-) method to remove a setting with the name "themeColor". The only parameter of the **remove** method is the case-sensitive _name_ of the setting.</span></span>


```js
Office.context.document.settings.remove('themeColor');
```

<span data-ttu-id="63333-182">該当する設定が存在しない場合は何も起きません。</span><span class="sxs-lookup"><span data-stu-id="63333-182">Nothing will happen if the setting does not exist.</span></span> <span data-ttu-id="63333-183">ドキュメントから設定を削除したままにする場合は、**Settings.saveAsync** メソッドを使用します。</span><span class="sxs-lookup"><span data-stu-id="63333-183">Use the **Settings.saveAsync** method to persist removal of the setting from the document.</span></span>


### <a name="saving-your-settings"></a><span data-ttu-id="63333-184">設定の保存</span><span class="sxs-lookup"><span data-stu-id="63333-184">Saving your settings</span></span>

<span data-ttu-id="63333-p121">現在のセッション中に、アドインがメモリ内の設定プロパティ バッグに対して行った追加、変更、または削除を保存するには、[Settings.saveAsync](/javascript/api/office/office.settings#saveasync-options--callback-) メソッドを呼び出してそれらの設定をドキュメントに保存する必要があります。**saveAsync** メソッドの唯一のパラメーターは _callback_ であり、これはパラメーターを 1 つだけ取るコールバック関数です。</span><span class="sxs-lookup"><span data-stu-id="63333-p121">To save any additions, changes, or deletions your add-in made to the in-memory copy of the settings property bag during the current session, you must call the [Settings.saveAsync](/javascript/api/office/office.settings#saveasync-options--callback-) method to store them in the document. The only parameter of the **saveAsync** method is _callback_, which is a callback function with a single parameter.</span></span> 


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

<span data-ttu-id="63333-187">**saveAsync** メソッドに _callback_ パラメーターとして渡した匿名関数は、操作の完了時に実行されます。</span><span class="sxs-lookup"><span data-stu-id="63333-187">The anonymous function passed into the **saveAsync** method as the _callback_ parameter is executed when the operation is completed.</span></span> <span data-ttu-id="63333-188">コールバックの _asyncResult_ パラメーターは、処理の状況を含む **AsyncResult** オブジェクトへのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="63333-188">The _asyncResult_ parameter of the callback provides access to an **AsyncResult** object that contains the status of the operation.</span></span> <span data-ttu-id="63333-189">例では、関数が **AsyncResult.status** プロパティを調べて、保存操作が成功したのか失敗したのかを確認し、アドインのページにその結果を表示します。</span><span class="sxs-lookup"><span data-stu-id="63333-189">In the example, the function checks the **AsyncResult.status** property to see if the save operation succeeded or failed, and then displays the result in the add-in's page.</span></span>

## <a name="how-to-save-custom-xml-to-the-document"></a><span data-ttu-id="63333-190">ドキュメントにカスタム XML を保存する方法</span><span class="sxs-lookup"><span data-stu-id="63333-190">How to save custom XML to the document</span></span>

> [!NOTE]
> <span data-ttu-id="63333-p123">このセクションでは、Word でサポートされている Office 共通 JavaScript API のコンテキストでのカスタム XML 部分について説明します。 ホスト固有の Excel JavaScript API でも、カスタム XML 部分にアクセスできます。 Excel の API とプログラミング パターンには、わずかな違いがあります。 詳細については、[Excel の CustomXmlPart](/javascript/api/excel/excel.customxmlpart) を参照してください。</span><span class="sxs-lookup"><span data-stu-id="63333-p123">This section discusses custom XML parts in the context of the Office Common JavaScript API which is supported in Word. The host-specific Excel JavaScript API also provides access to the custom XML parts. The Excel APIs and programming patterns are somewhat different. For more information, see [Excel CustomXmlPart](/javascript/api/excel/excel.customxmlpart).</span></span>

<span data-ttu-id="63333-195">ドキュメントの Settings のサイズ制限を超過する情報や構造化された特徴を持つ情報を保存する必要がある場合には、追加のストレージ オプションがあります。</span><span class="sxs-lookup"><span data-stu-id="63333-195">There is an addtional storage option when you need to store information that exceeds the size limits of the document Settings or which has a structured character.</span></span> <span data-ttu-id="63333-196">Word および Excel の作業ウィンドウ アドインには、カスタムの XML マークアップを保持できます (Excel については、このセクションの冒頭にあるノートを参照してください)。</span><span class="sxs-lookup"><span data-stu-id="63333-196">You can persist custom XML markup in a task pane add-in for Word (and for Excel, but see the note at the top of this section).</span></span> <span data-ttu-id="63333-197">Word の場合は、[CustomXmlPart](/javascript/api/office/office.customxmlpart) とそのメソッドを使用します (繰り返しになりますが、Excel の場合は上記のノートを参照してください)。</span><span class="sxs-lookup"><span data-stu-id="63333-197">In Word, you use the [CustomXmlPart](/javascript/api/office/office.customxmlpart) object and its methods (again, see the note above for Excel).</span></span> <span data-ttu-id="63333-198">次のコードでは、カスタム XML パーツを作成して、その ID とコンテンツをページの div に表示します。</span><span class="sxs-lookup"><span data-stu-id="63333-198">The following code creates a custom XML part and displays its ID and then its content in divs on the page.</span></span> <span data-ttu-id="63333-199">XML 文字列には `xmlns` 属性が必ず存在する点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="63333-199">Note that there must be an `xmlns` attribute in the XML string.</span></span>

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

<span data-ttu-id="63333-p125">カスタム XML 部分を取得するには、[getByIdAsync](/javascript/api/office/office.customxmlparts#getbyidasync-id--options--callback-) メソッドを使用しますが、ID は XML 部分の作成時に生成された GUID になるため、コードの作成時に ID の内容を知ることはできません。 そのため、XML 部分を作成したら、その XML 部分の ID を設定としてすぐに保存して、覚えやすいキーを割り当てることがベスト プラクティスになります。 次のメソッドは、この方法を示してます  (ただし、カスタム設定の操作に関する詳細とベスト プラクティスについては、この記事の前半のセクションを参照してください)。</span><span class="sxs-lookup"><span data-stu-id="63333-p125">To retrieve a custom XML part, you use the [getByIdAsync](/javascript/api/office/office.customxmlparts#getbyidasync-id--options--callback-) method, but the ID is a GUID that is generated when the XML part is created, so you can't know when coding what the ID is. For that reason, it is a good practice when creating an XML part to immediately store the ID of the XML part as a setting and give it a memorable key. The following method shows how to do this. (But see earlier sections of this article for details and best practices when working with custom settings).</span></span>

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

<span data-ttu-id="63333-204">次のコードは、最初に設定から ID を取得することで、XML 部分を取得する方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="63333-204">The following code shows how to retrieve the XML part by first getting its ID from a setting.</span></span>

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


## <a name="how-to-save-settings-in-the-users-mailbox-for-outlook-add-ins-as-roaming-settings"></a><span data-ttu-id="63333-205">Outlook アドインでユーザーのメールボックスに設定をローミング設定として保存する方法</span><span class="sxs-lookup"><span data-stu-id="63333-205">How to save settings in the user's mailbox for Outlook add-ins as roaming settings</span></span>


<span data-ttu-id="63333-206">Outlook アドインは、[RoamingSettings](/javascript/api/outlook/office.roamingsettings) オブジェクトを使用して、ユーザーのメールボックスに固有の、アドインの状態および設定のデータを保存できます。</span><span class="sxs-lookup"><span data-stu-id="63333-206">An Outlook add-in can use the [RoamingSettings](/javascript/api/outlook/office.roamingsettings) object to save add-in state and settings data that is specific to the user's mailbox.</span></span> <span data-ttu-id="63333-207">このデータには、アドインを実行しているユーザーではなく、Outlook アドインのみがアクセスできます。</span><span class="sxs-lookup"><span data-stu-id="63333-207">This data is accessible only by that Outlook add-in on behalf of the user running the add-in.</span></span> <span data-ttu-id="63333-208">データはユーザーの Exchange Server メールボックスに格納されます。データには、ユーザーが自分のアカウントにログインして Outlook アドインを実行したときにアクセスできるようになります。</span><span class="sxs-lookup"><span data-stu-id="63333-208">The data is stored on the user's Exchange Server mailbox, and is accessible when that user logs into their account and runs the Outlook add-in.</span></span>


### <a name="loading-roaming-settings"></a><span data-ttu-id="63333-209">ローミング設定の読み込み</span><span class="sxs-lookup"><span data-stu-id="63333-209">Loading roaming settings</span></span>


<span data-ttu-id="63333-p127">通常、Outlook アドインでは、 [Office.initialize](/javascript/api/office) イベント ハンドラーでローミング設定を読み込みます。次の JavaScript のコード例は、既存のローミング設定を読み込む方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="63333-p127">An Outlook add-in typically loads roaming settings in the [Office.initialize](/javascript/api/office) event handler. The following JavaScript code example shows how to load existing roaming settings.</span></span>


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


### <a name="creating-or-assigning-a-roaming-setting"></a><span data-ttu-id="63333-212">ローミング設定の作成または割り当て</span><span class="sxs-lookup"><span data-stu-id="63333-212">Creating or assigning a roaming setting</span></span>


<span data-ttu-id="63333-p128">前の例に続けて、次の  `setAppSetting` 関数では、 [RoamingSettings.set](/javascript/api/outlook/office.roamingsettings#set-name--value-) メソッドを使用して、 `cookie` という名前の設定項目に今日の日付を設定、または今日の日付で更新する方法を示しています。次に、 [RoamingSettings.saveAsync](/javascript/api/outlook/office.roamingsettings#saveasync-callback-) メソッドを使用して Exchange Server にすべてのローミング設定を保存し直しています。</span><span class="sxs-lookup"><span data-stu-id="63333-p128">Continuing with the preceding example, the following  `setAppSetting` function shows how to use the [RoamingSettings.set](/javascript/api/outlook/office.roamingsettings#set-name--value-) method to set or update a setting named `cookie` with today's date. Then, it saves all the roaming settings back to the Exchange Server with the [RoamingSettings.saveAsync](/javascript/api/outlook/office.roamingsettings#saveasync-callback-) method.</span></span>


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

<span data-ttu-id="63333-215">**saveAsync** メソッドは、ローミング設定を非同期で保存し、オプションのコールバック関数を受け取ります。</span><span class="sxs-lookup"><span data-stu-id="63333-215">The **saveAsync** method saves roaming settings asynchronously and takes an optional callback function.</span></span> <span data-ttu-id="63333-216">このコード例では、`saveMyAppSettingsCallback` という名前のコールバック関数を **saveAsync** メソッドに渡します。</span><span class="sxs-lookup"><span data-stu-id="63333-216">This code sample passes a callback function named `saveMyAppSettingsCallback` to the **saveAsync** method.</span></span> <span data-ttu-id="63333-217">非同期呼び出しが返されると、`saveMyAppSettingsCallback` 関数の _asyncResult_ パラメーターが [AsyncResult](/javascript/api/outlook) オブジェクトにアクセスします。このオブジェクトを使用すると、**AsyncResult.status** プロパティで操作の成功または失敗を判定することができます。</span><span class="sxs-lookup"><span data-stu-id="63333-217">When the asynchronous call returns, the _asyncResult_ parameter of the `saveMyAppSettingsCallback` function provides access to an [AsyncResult](/javascript/api/outlook) object that you can use to determine the success or failure of the operation with the **AsyncResult.status** property.</span></span>


### <a name="removing-a-roaming-setting"></a><span data-ttu-id="63333-218">ローミング設定の削除</span><span class="sxs-lookup"><span data-stu-id="63333-218">Removing a roaming setting</span></span>


<span data-ttu-id="63333-219">また、次の  `removeAppSetting` 関数は、前の例をさらに拡張するものです。この例では、 [RoamingSettings.remove](/javascript/api/outlook/office.roamingsettings#remove-name-) メソッドを使用して `cookie` 設定を削除し、すべてのローミング設定を Exchange Server に保存し直す方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="63333-219">Also extending the preceding examples, the following  `removeAppSetting` function, shows how to use the [RoamingSettings.remove](/javascript/api/outlook/office.roamingsettings#remove-name-) method to remove the `cookie` setting and save all the roaming settings back to the Exchange Server.</span></span>


```js
// Remove an application setting.
function removeAppSetting()
{
    _settings.remove("cookie");
    _settings.saveAsync(saveMyAppSettingsCallback);
}
```


## <a name="how-to-save-settings-per-item-for-outlook-add-ins-as-custom-properties"></a><span data-ttu-id="63333-220">Outlook アドインでアイテムごとに設定をカスタムプロパティとして保存する方法</span><span class="sxs-lookup"><span data-stu-id="63333-220">How to save settings per item for Outlook add-ins as custom properties</span></span>


<span data-ttu-id="63333-p130">カスタム プロパティを使用すると、Outlook アドインは処理しているアイテムに関する情報を保存できます。たとえば、Outlook アドインを使用して、メッセージ内の会議の提案から予定を作成する場合は、カスタム プロパティを使用して、会議が作成されたという事実を保存できます。これにより、メッセージを再び開いたときに、Outlook アドインが再び予定の作成を行うことはありません。</span><span class="sxs-lookup"><span data-stu-id="63333-p130">Custom properties let your Outlook add-in store information about an item it is working with. For example, if your Outlook add-in creates an appointment from a meeting suggestion in a message, you can use custom properties to store the fact that the meeting was created. This makes sure that if the message is opened again, your Outlook add-in doesn't offer to create the appointment again.</span></span>

<span data-ttu-id="63333-p131">メッセージ、予定、または会議出席依頼の特定のアイテムに対してカスタム プロパティを使用するには、その前に、 [Item](/javascript/api/outlook/office.mailbox) オブジェクトの **loadCustomPropertiesAsync** メソッドを呼び出して、プロパティをメモリに読み込む必要があります。現在のアイテムに対してカスタム プロパティが既に設定されている場合は、この時点で Exchange サーバーから読み込まれます。プロパティを読み込んだ後、 [CustomProperties](/javascript/api/outlook/office.customproperties#set-name--value-) オブジェクトの [set](/javascript/api/outlook/office.roamingsettings) メソッドおよび **get** メソッドを使用して、メモリ内のプロパティの追加、更新、および取得を実行できます。アイテムのカスタム プロパティに対して行った変更を保存するには、 [saveAsync](/javascript/api/outlook/office.customproperties#saveasync-callback--asynccontext-) メソッドを使用して、アイテムに加えた変更を Exchange サーバー上で保持する必要があります。</span><span class="sxs-lookup"><span data-stu-id="63333-p131">Before you can use custom properties for a particular message, appointment, or meeting request item, you must load the properties into memory by calling the [loadCustomPropertiesAsync](/javascript/api/outlook/office.mailbox) method of the **Item** object. If any custom properties are already set for the current item, they are loaded from the Exchange server at this point. After you have loaded the properties, you can use the [set](/javascript/api/outlook/office.customproperties#set-name--value-) and [get](/javascript/api/outlook/office.roamingsettings) methods of the **CustomProperties** object to add, update, and retrieve properties in memory. To save any changes that you make to the item's custom properties, you must use the [saveAsync](/javascript/api/outlook/office.customproperties#saveasync-callback--asynccontext-) method to persist the changes to the item on the Exchange server.</span></span>


### <a name="custom-properties-example"></a><span data-ttu-id="63333-228">カスタム プロパティの例</span><span class="sxs-lookup"><span data-stu-id="63333-228">Custom properties example</span></span>

<span data-ttu-id="63333-p132">以下の例では、カスタム プロパティを使用する Outlook アドインの一連の関数を、簡略化して示しています。この例を出発点として、カスタム プロパティを使用する Outlook アドインを作成できます。</span><span class="sxs-lookup"><span data-stu-id="63333-p132">The following example shows a simplified set of functions for an Outlook add-in that uses custom properties. You can use this example as a starting point for your Outlook add-in that uses custom properties.</span></span> 

<span data-ttu-id="63333-231">これらの関数を使用する Outlook アドインは、次の例に示すように、`_customProps` 変数で **get** メソッドを呼び出すことによって、任意のカスタム プロパティを取得します。</span><span class="sxs-lookup"><span data-stu-id="63333-231">An Outlook add-in that uses these functions retrieves any custom properties by calling the **get** method on the `_customProps` variable, as shown in the following example.</span></span>




```js
var property = _customProps.get("propertyName");
```

<span data-ttu-id="63333-232">以下の例には、次の関数が含まれています。</span><span class="sxs-lookup"><span data-stu-id="63333-232">This example includes the following functions:</span></span>



|<span data-ttu-id="63333-233">**関数名**</span><span class="sxs-lookup"><span data-stu-id="63333-233">**Function name**</span></span>|<span data-ttu-id="63333-234">**説明**</span><span class="sxs-lookup"><span data-stu-id="63333-234">**Description**</span></span>|
|:-----|:-----|
| `Office.initialize`|<span data-ttu-id="63333-235">アドインを初期化し、Exchange サーバーから現在のアイテムのカスタム プロパティを読み込みます。</span><span class="sxs-lookup"><span data-stu-id="63333-235">Initializes the add-in and loads the custom properties for the current item from the Exchange server.</span></span>|
| `customPropsCallback`|<span data-ttu-id="63333-236">Exchange サーバーから返されるカスタム プロパティを取得し、後で使用できるように保存します。</span><span class="sxs-lookup"><span data-stu-id="63333-236">Gets the custom properties that are returned from the Exchange server and saves it for later use.</span></span>|
| `updateProperty`|<span data-ttu-id="63333-237">特定のプロパティを設定または更新し、その変更を Exchange サーバーに保存します。</span><span class="sxs-lookup"><span data-stu-id="63333-237">Sets or updates a specific property, and then saves the change to the Exchange server.</span></span>|
| `removeProperty`|<span data-ttu-id="63333-238">特定のプロパティを削除し、その削除を Exchange サーバーに保存します。</span><span class="sxs-lookup"><span data-stu-id="63333-238">Removes a specific property, and then persists the removal to the Exchange server.</span></span>|
| `saveCallback`|<span data-ttu-id="63333-239">`updateProperty` 関数および `removeProperty` 関数内で **saveAsync** メソッドを呼び出すためのコールバック。</span><span class="sxs-lookup"><span data-stu-id="63333-239">Callback for calls to the **saveAsync** method in the `updateProperty` and `removeProperty` functions.</span></span>|



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


## <a name="see-also"></a><span data-ttu-id="63333-240">関連項目</span><span class="sxs-lookup"><span data-stu-id="63333-240">See also</span></span>

- [<span data-ttu-id="63333-241">JavaScript API for Office について</span><span class="sxs-lookup"><span data-stu-id="63333-241">Understanding the JavaScript API for Office</span></span>](understanding-the-javascript-api-for-office.md)
- [<span data-ttu-id="63333-242">Outlook アドイン</span><span class="sxs-lookup"><span data-stu-id="63333-242">Outlook add-ins</span></span>](/outlook/add-ins/)
- [<span data-ttu-id="63333-243">Excel-Add-in-JavaScript-PersistCustomSettings</span><span class="sxs-lookup"><span data-stu-id="63333-243">Excel-Add-in-JavaScript-PersistCustomSettings</span></span>](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings)
