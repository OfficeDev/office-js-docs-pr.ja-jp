---
title: アドインの状態および設定を保持する
description: ブラウザー コントロールのステートレスOfficeで実行されているアドイン Web アプリケーションにデータを保持する方法について説明します。
ms.date: 03/23/2021
localization_priority: Normal
ms.openlocfilehash: 47f1b9a5000660f13f8bba5e747d0f5ca5e9d6a0
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/09/2021
ms.locfileid: "51651080"
---
# <a name="persisting-add-in-state-and-settings"></a><span data-ttu-id="843d5-103">アドインの状態および設定を保持する</span><span class="sxs-lookup"><span data-stu-id="843d5-103">Persisting add-in state and settings</span></span>

[!include[information about the common API](../includes/alert-common-api-info.md)]

<span data-ttu-id="843d5-p101">Office アドインは、基本的にブラウザー コントロールのステートレス環境で動作する Web アプリケーションです。したがって、アドインでは、そのアドインを使用するセッション間で特定の操作または機能を継続して維持するためのデータを保持することが必要な場合があります。たとえば、アドインには、ユーザーの優先ビューや既定の場所など、アドインで保存しておき、アドインが次回初期化されたときにリロードする必要があるカスタム設定または他の値が含まれる場合があります。その場合は、次のようにします。</span><span class="sxs-lookup"><span data-stu-id="843d5-p101">Office Add-ins are essentially web applications running in the stateless environment of a browser control. As a result, your add-in may need to persist data to maintain the continuity of certain operations or features across sessions of using your add-in. For example, your add-in may have custom settings or other values that it needs to save and reload the next time it's initialized, such as a user's preferred view or default location. To do that, you can:</span></span>

- <span data-ttu-id="843d5-108">データを格納する javaScript API Officeメンバーを次のように使用します。</span><span class="sxs-lookup"><span data-stu-id="843d5-108">Use members of the Office JavaScript API that store data as either:</span></span>
  - <span data-ttu-id="843d5-109">アドインの種類応じた場所に保存されるプロパティ バッグ内の名前と値の組。</span><span class="sxs-lookup"><span data-stu-id="843d5-109">Name/value pairs in a property bag stored in a location that depends on add-in type.</span></span>
  - <span data-ttu-id="843d5-110">ドキュメント内に保存されるカスタム XML。</span><span class="sxs-lookup"><span data-stu-id="843d5-110">Custom XML stored in the document.</span></span>

- <span data-ttu-id="843d5-111">基になるブラウザー コントロールによって提供される技術である、ブラウザーの Cookie、または HTML5 Web ストレージ ([localStorage](https://developer.mozilla.org/docs/Web/API/Window/localStorage) または [sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage)) を使用します。</span><span class="sxs-lookup"><span data-stu-id="843d5-111">Use techniques provided by the underlying browser control: browser cookies, or HTML5 web storage ([localStorage](https://developer.mozilla.org/docs/Web/API/Window/localStorage) or [sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage)).</span></span>
    > [!NOTE]
    > <span data-ttu-id="843d5-112">ユーザーは、選択した設定に応じて、ブラウザー ベースのストレージ手法をブロックできます。</span><span class="sxs-lookup"><span data-stu-id="843d5-112">The user can block browser-based storage techniques depending on the settings they choose.</span></span>

<span data-ttu-id="843d5-113">この記事では、JavaScript API Officeを使用して現在のドキュメントにアドインの状態を保持する方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="843d5-113">This article focuses on how to use the Office JavaScript API to persist add-in state to the current document.</span></span> <span data-ttu-id="843d5-114">開いているドキュメントのユーザー設定の追跡など、ドキュメント間で状態を保持する必要がある場合は、別の方法を使用する必要があります。</span><span class="sxs-lookup"><span data-stu-id="843d5-114">If you need to persist state across documents, such as tracking user preferences across any documents they open, you will need to use a different approach.</span></span> <span data-ttu-id="843d5-115">たとえば [、SSO](sso-in-office-add-ins.md#using-the-sso-token-as-an-identity) を使用してユーザー ID を取得し、ユーザー ID とその設定をオンライン データベースに保存できます。</span><span class="sxs-lookup"><span data-stu-id="843d5-115">For example, you could use [SSO](sso-in-office-add-ins.md#using-the-sso-token-as-an-identity) to obtain the user identity, and then save the user ID and their settings to an online database.</span></span>

## <a name="persisting-add-in-state-and-settings-with-the-office-javascript-api"></a><span data-ttu-id="843d5-116">JavaScript API を使用してアドインの状態と設定Office保持する</span><span class="sxs-lookup"><span data-stu-id="843d5-116">Persisting add-in state and settings with the Office JavaScript API</span></span>

<span data-ttu-id="843d5-117">JavaScript API Officeには、次の[](/javascript/api/office/office.settings)表で説明するように、セッション間でアドインの状態を保存する Settings オブジェクト[、RoamingSettings](/javascript/api/outlook/office.roamingsettings)オブジェクト、[および CustomProperties](/javascript/api/outlook/office.customproperties)オブジェクトが提供されています。</span><span class="sxs-lookup"><span data-stu-id="843d5-117">The Office JavaScript API provides the [Settings](/javascript/api/office/office.settings), [RoamingSettings](/javascript/api/outlook/office.roamingsettings), and [CustomProperties](/javascript/api/outlook/office.customproperties) objects for saving add-in state across sessions as described in the following table.</span></span> <span data-ttu-id="843d5-118">すべてのケースで、保存された設定値は、それを作成したアドインの [Id](../reference/manifest/id.md) にのみ関連付けられます。</span><span class="sxs-lookup"><span data-stu-id="843d5-118">In all cases, the saved settings values are associated with the [Id](../reference/manifest/id.md) of the add-in that created them.</span></span>

|<span data-ttu-id="843d5-119">**オブジェクト**</span><span class="sxs-lookup"><span data-stu-id="843d5-119">**Object**</span></span>|<span data-ttu-id="843d5-120">**アドインの種類のサポート**</span><span class="sxs-lookup"><span data-stu-id="843d5-120">**Add-in type support**</span></span>|<span data-ttu-id="843d5-121">**ストレージの場所**</span><span class="sxs-lookup"><span data-stu-id="843d5-121">**Storage location**</span></span>|<span data-ttu-id="843d5-122">**Officeサポート**</span><span class="sxs-lookup"><span data-stu-id="843d5-122">**Office application support**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="843d5-123">Settings</span><span class="sxs-lookup"><span data-stu-id="843d5-123">Settings</span></span>](/javascript/api/office/office.settings)|<span data-ttu-id="843d5-124">コンテンツおよび作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="843d5-124">content and task pane</span></span>|<span data-ttu-id="843d5-125">アドインが連携しているドキュメント、スプレッドシート、またはプレゼンテーション。</span><span class="sxs-lookup"><span data-stu-id="843d5-125">The document, spreadsheet, or presentation the add-in is working with.</span></span> <span data-ttu-id="843d5-126">コンテンツおよび作業ウィンドウのアドイン設定は、その設定が保存されているドキュメントから、その設定を作成したアドインで使用できます。</span><span class="sxs-lookup"><span data-stu-id="843d5-126">Content and task pane add-in settings are available to the add-in that created them from the document where they are saved.</span></span><br/><br/><span data-ttu-id="843d5-p105">**重要:\*\*\*\*Settings** オブジェクトを使用して、パスワードおよびその他の機密の個人を特定できる情報 (PII) を保存しないでください。保存されたデータはユーザーに対して表示されませんが、ドキュメントの一部として保存されているため、ドキュメントのファイル形式を直接読み取ることでアクセスできます。アドインによる PII の使用と、アドインが必要とするすべての PII の保存は、開発するアドインをユーザーのセキュリティが保護されるリソースとしてホストするサーバーのみで行うよう制限する必要があります。</span><span class="sxs-lookup"><span data-stu-id="843d5-p105">**Important:** Don't store passwords and other sensitive personally identifiable information (PII) with the **Settings** object. The data saved isn't visible to end users, but it is stored as part of the document, which is accessible by reading the document's file format directly. You should limit your add-in's use of PII and store any PII required by your add-in only on the server hosting your add-in as a user-secured resource.</span></span>|<span data-ttu-id="843d5-130">Word、Excel、または PowerPoint</span><span class="sxs-lookup"><span data-stu-id="843d5-130">Word, Excel, or PowerPoint</span></span><br/><br/> <span data-ttu-id="843d5-131">**メモ:** Project 2013 の作業ウィンドウ アドインでは、アドインの状態または設定を保存するための **Settings** API をサポートしていません。</span><span class="sxs-lookup"><span data-stu-id="843d5-131">**Note:** Task pane add-ins for Project 2013 don't support the **Settings** API for storing add-in state or settings.</span></span> <span data-ttu-id="843d5-132">ただし、Project で実行されているアドイン (および他の Office クライアント アプリケーション) では、ブラウザー Cookie や Web ストレージなどの手法を使用できます。</span><span class="sxs-lookup"><span data-stu-id="843d5-132">However, for add-ins running in Project (as well as other Office client applications) you can use techniques such as browser cookies or web storage.</span></span> <span data-ttu-id="843d5-133">こうした技術の詳細については、「[Excel-Add-in-JavaScript-PersistCustomSettings](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="843d5-133">For more information on these techniques, see the [Excel-Add-in-JavaScript-PersistCustomSettings](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings).</span></span> |
|[<span data-ttu-id="843d5-134">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="843d5-134">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings)|<span data-ttu-id="843d5-135">Outlook</span><span class="sxs-lookup"><span data-stu-id="843d5-135">Outlook</span></span>|<span data-ttu-id="843d5-136">アドインがインストールされている、ユーザーの Exchange サーバー メールボックス。</span><span class="sxs-lookup"><span data-stu-id="843d5-136">The user's Exchange server mailbox where the add-in is installed.</span></span> <span data-ttu-id="843d5-137">これらの設定はユーザーのサーバー メールボックスに格納されますので、ユーザーと一緒に "ローミング" し、サポートされている Office クライアント アプリケーションまたはブラウザーがユーザーのメールボックスにアクセスするコンテキストでアドインを実行している場合に使用できます。</span><span class="sxs-lookup"><span data-stu-id="843d5-137">Because these settings are stored in the user's server mailbox, they can "roam" with the user and are available to the add-in when it is running in the context of any supported Office client application or browser accessing that user's mailbox.</span></span><br/><br/> <span data-ttu-id="843d5-138">Outlook アドインのローミング設定は、その設定を作成したアドインのみが利用でき、また、アドインがインストールされているメールボックスからのみ利用できます。</span><span class="sxs-lookup"><span data-stu-id="843d5-138">Outlook add-in roaming settings are available only to the add-in that created them, and only from the mailbox where the add-in is installed.</span></span>|<span data-ttu-id="843d5-139">Outlook</span><span class="sxs-lookup"><span data-stu-id="843d5-139">Outlook</span></span>|
|[<span data-ttu-id="843d5-140">CustomProperties</span><span class="sxs-lookup"><span data-stu-id="843d5-140">CustomProperties</span></span>](/javascript/api/outlook/office.customproperties)|<span data-ttu-id="843d5-141">Outlook</span><span class="sxs-lookup"><span data-stu-id="843d5-141">Outlook</span></span>|<span data-ttu-id="843d5-p108">アドインが連携するメッセージ、予定、または会議出席依頼アイテム。 Outlook アドイン アイテムのカスタム プロパティは、そのプロパティを作成したアドインのみが利用でき、また、プロパティが保存されているアイテムからのみ利用できます。</span><span class="sxs-lookup"><span data-stu-id="843d5-p108">The message, appointment, or meeting request item the add-in is working with. Outlook add-in item custom properties are available only to the add-in that created them, and only from the item where they are saved.</span></span>|<span data-ttu-id="843d5-144">Outlook</span><span class="sxs-lookup"><span data-stu-id="843d5-144">Outlook</span></span>|
|[<span data-ttu-id="843d5-145">CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="843d5-145">CustomXmlParts</span></span>](/javascript/api/office/office.customxmlparts)|<span data-ttu-id="843d5-146">作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="843d5-146">task pane</span></span>|<span data-ttu-id="843d5-p109">アドインが連携しているドキュメント、スプレッドシート、またはプレゼンテーション。作業ウィンドウのアドイン設定は、その設定が保存されているドキュメントから、その設定を作成したアドインで使用できます。</span><span class="sxs-lookup"><span data-stu-id="843d5-p109">The document, spreadsheet, or presentation the add-in is working with. Task pane add-in settings are available to the add-in that created them from the document where they are saved.</span></span><br/><br/><span data-ttu-id="843d5-p110">**重要:** カスタム XML 部分には、パスワードなどの個人情報 (PII) を保存しないでください。保存されたデータはユーザーに対して表示されませんが、ドキュメントの一部として保存されるため、ドキュメントのファイル形式を直接読み取ることでアクセスできます。アドインによる PII の使用と、アドインが必要とするすべての PII の保存は、開発するアドインをユーザーのセキュリティが保護されるリソースとしてホストするサーバーのみで行うよう制限する必要があります。</span><span class="sxs-lookup"><span data-stu-id="843d5-p110">**Important:** Don't store passwords and other sensitive personally identifiable information (PII) in a custom XML part. The data saved isn't visible to end users, but it is stored as part of the document, which is accessible by reading the document's file format directly. You should limit your add-in's use of PII and store any PII required by your add-in only on the server hosting your add-in as a user-secured resource.</span></span>|<span data-ttu-id="843d5-152">Word (JavaScript 共通 API Office使用) Excel (アプリケーション固有の Excel JavaScript API を使用)</span><span class="sxs-lookup"><span data-stu-id="843d5-152">Word (using the Office JavaScript Common API) Excel (using the application-specific Excel JavaScript API</span></span>|

## <a name="settings-data-is-managed-in-memory-at-runtime"></a><span data-ttu-id="843d5-153">実行時のメモリ内での設定データの管理</span><span class="sxs-lookup"><span data-stu-id="843d5-153">Settings data is managed in memory at runtime</span></span>

> [!NOTE]
> <span data-ttu-id="843d5-154">この後の 2 つのセクションでは、Office 共通 JavaScript API のコンテキストでの設定について説明します。</span><span class="sxs-lookup"><span data-stu-id="843d5-154">The following two sections discuss settings in the context of the Office Common JavaScript API.</span></span> <span data-ttu-id="843d5-155">アプリケーション固有の Excel JavaScript API では、カスタム設定にもアクセスできます。</span><span class="sxs-lookup"><span data-stu-id="843d5-155">The application-specific Excel JavaScript API also provides access to the custom settings.</span></span> <span data-ttu-id="843d5-156">Excel の API とプログラミング パターンには、わずかな違いがあります。</span><span class="sxs-lookup"><span data-stu-id="843d5-156">The Excel APIs and programming patterns are somewhat different.</span></span> <span data-ttu-id="843d5-157">詳細については、[Excel の SettingCollection](/javascript/api/excel/excel.settingcollection) を参照してください。</span><span class="sxs-lookup"><span data-stu-id="843d5-157">For more information, see [Excel SettingCollection](/javascript/api/excel/excel.settingcollection).</span></span>

<span data-ttu-id="843d5-158">内部的には、 、 オブジェクトでアクセスされるプロパティ バッグ内のデータは、名前と値のペアを含むシリアル化 `Settings` `CustomProperties` された JavaScript オブジェクト表記 (JSON) オブジェクトとして `RoamingSettings` 格納されます。</span><span class="sxs-lookup"><span data-stu-id="843d5-158">Internally, the data in the property bag accessed with the `Settings`, `CustomProperties`, or `RoamingSettings` objects is stored as a serialized JavaScript Object Notation (JSON) object that contains name/value pairs.</span></span> <span data-ttu-id="843d5-159">各値の名前 (キー) は、 で、格納されている値は JavaScript 、関数である `string` `string` `number` `date` `object` **必要があります**。</span><span class="sxs-lookup"><span data-stu-id="843d5-159">The name (key) for each value must be a `string`, and the stored value can be a JavaScript `string`, `number`, `date`, or `object`, but not a **function**.</span></span>

<span data-ttu-id="843d5-160">この例はプロパティ バッグの構造を示し、3 つの定義された **string** 値 (`firstName`、`location`、`defaultView` という名前) が含まれます。</span><span class="sxs-lookup"><span data-stu-id="843d5-160">This example of the property bag structure contains three defined **string** values named `firstName`,  `location`, and  `defaultView`.</span></span>

```json
{
    "firstName":"Erik",
    "location":"98052",
    "defaultView":"basic"
}
```

<span data-ttu-id="843d5-161">設定プロパティ バッグは、前のアドイン セッション中に保存された後、アドインが初期化されるとき、またはその後はいつでも、アドインの現行セッション中は読み込むことができます。</span><span class="sxs-lookup"><span data-stu-id="843d5-161">After the settings property bag is saved during the previous add-in session, it can be loaded when the add-in is initialized or at any point after that during the add-in's current session.</span></span> <span data-ttu-id="843d5-162">セッション中、設定は、作成する設定の種類 `get` `set` (設定 `remove` **、CustomProperties、\*\*\*\*または RoamingSettings)** に対応するオブジェクトのメソッドを使用して、メモリ内で完全に管理されます。</span><span class="sxs-lookup"><span data-stu-id="843d5-162">During the session, the settings are managed in entirely in memory using the `get`, `set`, and `remove` methods of the object that corresponds to the kind of settings you are creating (**Settings**, **CustomProperties**, or **RoamingSettings**).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="843d5-163">アドインの現在のセッション中に行われた追加、更新、または削除を保存場所に保持するには、その種類の設定を処理するために使用する対応するオブジェクトのメソッドを呼び出す必要があります `saveAsync` 。</span><span class="sxs-lookup"><span data-stu-id="843d5-163">To persist any additions, updates, or deletions made during the add-in's current session to the storage location, you must call the `saveAsync` method of the corresponding object used to work with that kind of settings.</span></span> <span data-ttu-id="843d5-164">、およびメソッドは、settings プロパティ バッグのメモリ内 `get` `set` `remove` コピーでのみ動作します。</span><span class="sxs-lookup"><span data-stu-id="843d5-164">The `get`, `set`, and `remove` methods operate only on the in-memory copy of the settings property bag.</span></span> <span data-ttu-id="843d5-165">アドインを呼び出さずに閉じた場合、そのセッション中に設定に加えた変更 `saveAsync` は失われます。</span><span class="sxs-lookup"><span data-stu-id="843d5-165">If your add-in is closed without calling `saveAsync`, any changes made to settings during that session will be lost.</span></span>

## <a name="how-to-save-add-in-state-and-settings-per-document-for-content-and-task-pane-add-ins"></a><span data-ttu-id="843d5-166">コンテンツ アドインおよび作業ウィンドウ アドインで、ドキュメントごとにアドインの状態と設定を保存する方法</span><span class="sxs-lookup"><span data-stu-id="843d5-166">How to save add-in state and settings per document for content and task pane add-ins</span></span>

<span data-ttu-id="843d5-167">Word、Excel、または PowerPoint 用のコンテンツ アドインまたは作業ウィンドウ アドインの状態またはカスタム設定を保持するには、 [Settings](/javascript/api/office/office.settings) オブジェクトとそのメソッドを使用します。</span><span class="sxs-lookup"><span data-stu-id="843d5-167">To persist state or custom settings of a content or task pane add-in for Word, Excel, or PowerPoint, you use the [Settings](/javascript/api/office/office.settings) object and its methods.</span></span> <span data-ttu-id="843d5-168">オブジェクトのメソッドで作成されたプロパティ バッグは、そのオブジェクトを作成したコンテンツ アドインまたは作業ウィンドウ アドインのインスタンスでのみ使用できます。保存されているドキュメントからのみ使用 `Settings` できます。</span><span class="sxs-lookup"><span data-stu-id="843d5-168">The property bag created with the methods of the `Settings` object are available only to the instance of the content or task pane add-in that created it, and only from the document in which it is saved.</span></span>

<span data-ttu-id="843d5-169">オブジェクトは Document オブジェクトの一部として自動的に読み込まれ、作業ウィンドウアドインまたはコンテンツ アドインがアクティブ化されると `Settings` 使用できます。 [](/javascript/api/office/office.document)</span><span class="sxs-lookup"><span data-stu-id="843d5-169">The `Settings` object is automatically loaded as part of the [Document](/javascript/api/office/office.document) object, and is available when the task pane or content add-in is activated.</span></span> <span data-ttu-id="843d5-170">オブジェクトを `Document` インスタンス化した後、オブジェクトの settings プロパティ `Settings` を[](/javascript/api/office/office.document#settings)使用してオブジェクトにアクセス `Document` できます。</span><span class="sxs-lookup"><span data-stu-id="843d5-170">After the `Document` object is instantiated, you can access the `Settings` object with the [settings](/javascript/api/office/office.document#settings) property of the `Document` object.</span></span> <span data-ttu-id="843d5-171">セッションの有効期間中に、プロパティ バッグのメモリ内コピーから永続化された設定とアドインの状態を読み取り、書き込み、または削除するには、 およびメソッドを使用できます `Settings.get` `Settings.set` `Settings.remove` 。</span><span class="sxs-lookup"><span data-stu-id="843d5-171">During the lifetime of the session, you can just use the `Settings.get`, `Settings.set`, and `Settings.remove` methods to read, write, or remove persisted settings and add-in state from the in-memory copy of the property bag.</span></span>

<span data-ttu-id="843d5-172">set メソッドと remove メソッドは設定プロパティ バッグのメモリ内コピーに対してのみ動作するので、アドインが関連付けられているドキュメントに新しい設定を保存、または変更された設定を保存し直すには [Settings.saveAsync](/javascript/api/office/office.settings#saveasync-options--callback-) メソッドを呼び出す必要があります。</span><span class="sxs-lookup"><span data-stu-id="843d5-172">Because the set and remove methods operate against only the in-memory copy of the settings property bag, to save new or changed settings back to the document the add-in is associated with you must call the [Settings.saveAsync](/javascript/api/office/office.settings#saveasync-options--callback-) method.</span></span>

### <a name="creating-or-updating-a-setting-value"></a><span data-ttu-id="843d5-173">設定値の作成または更新</span><span class="sxs-lookup"><span data-stu-id="843d5-173">Creating or updating a setting value</span></span>

<span data-ttu-id="843d5-p117">次のコード例では、[Settings.set](/javascript/api/office/office.settings#set-name--value-) メソッドを使用して `'themeColor'` という名前の設定を作成し、値 `'green'` を指定する方法を説明します。set メソッドの最初のパラメーターは、設定するか作成する設定の _name_ (Id) であり、これは大文字と小文字が区別されます。2 番目のパラメーターは、設定の _value_ です。</span><span class="sxs-lookup"><span data-stu-id="843d5-p117">The following code example shows how to use the [Settings.set](/javascript/api/office/office.settings#set-name--value-) method to create a setting called `'themeColor'` with a value `'green'`. The first parameter of the set method is the case-sensitive  _name_ (Id) of the setting to set or create. The second parameter is the _value_ of the setting.</span></span>

```js
Office.context.document.settings.set('themeColor', 'green');
```

 <span data-ttu-id="843d5-177">指定した名前を持つ設定は、それがまだ存在していない場合には作成され、すでに存在している場合はその値が更新されます。</span><span class="sxs-lookup"><span data-stu-id="843d5-177">The setting with the specified name is created if it doesn't already exist, or its value is updated if it does exist.</span></span> <span data-ttu-id="843d5-178">新しい `Settings.saveAsync` 設定または更新された設定をドキュメントに保持するには、このメソッドを使用します。</span><span class="sxs-lookup"><span data-stu-id="843d5-178">Use the `Settings.saveAsync` method to persist the new or updated settings to the document.</span></span>

### <a name="getting-the-value-of-a-setting"></a><span data-ttu-id="843d5-179">設定値の取得</span><span class="sxs-lookup"><span data-stu-id="843d5-179">Getting the value of a setting</span></span>

<span data-ttu-id="843d5-180">次の例では、 [Settings.get](/javascript/api/office/office.settings#get-name-) メソッドを使用して "themeColor" という名前の設定値を取得する方法を示します。</span><span class="sxs-lookup"><span data-stu-id="843d5-180">The following example shows how use the [Settings.get](/javascript/api/office/office.settings#get-name-) method to get the value of a setting called "themeColor".</span></span> <span data-ttu-id="843d5-181">メソッドの唯一の `get` パラメーターは、設定の大文字と小文字を区別する名前です。</span><span class="sxs-lookup"><span data-stu-id="843d5-181">The only parameter of the `get` method is the case-sensitive _name_ of the setting.</span></span>

```js
write('Current value for mySetting: ' + Office.context.document.settings.get('themeColor'));

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message;
}
```

 <span data-ttu-id="843d5-182">メソッド `get` は、渡された設定名に対して以前に保存された値を返します。</span><span class="sxs-lookup"><span data-stu-id="843d5-182">The `get` method returns the value that was previously saved for the setting _name_ that was passed in.</span></span> <span data-ttu-id="843d5-183">設定が存在しない場合、メソッドは **null** を返します。</span><span class="sxs-lookup"><span data-stu-id="843d5-183">If the setting doesn't exist, the method returns **null**.</span></span>

### <a name="removing-a-setting"></a><span data-ttu-id="843d5-184">設定の削除</span><span class="sxs-lookup"><span data-stu-id="843d5-184">Removing a setting</span></span>

<span data-ttu-id="843d5-185">次の例では、 [Settings.remove](/javascript/api/office/office.settings#remove-name-) メソッドを使用して、"themeColor" という名前の設定を削除する方法を示します。</span><span class="sxs-lookup"><span data-stu-id="843d5-185">The following example shows how to use the [Settings.remove](/javascript/api/office/office.settings#remove-name-) method to remove a setting with the name "themeColor".</span></span> <span data-ttu-id="843d5-186">メソッドの唯一の `remove` パラメーターは、設定の大文字と小文字を区別する名前です。</span><span class="sxs-lookup"><span data-stu-id="843d5-186">The only parameter of the `remove` method is the case-sensitive _name_ of the setting.</span></span>

```js
Office.context.document.settings.remove('themeColor');
```

<span data-ttu-id="843d5-187">該当する設定が存在しない場合は何も起きません。</span><span class="sxs-lookup"><span data-stu-id="843d5-187">Nothing will happen if the setting does not exist.</span></span> <span data-ttu-id="843d5-188">ドキュメントから `Settings.saveAsync` 設定の削除を保持するには、このメソッドを使用します。</span><span class="sxs-lookup"><span data-stu-id="843d5-188">Use the `Settings.saveAsync` method to persist removal of the setting from the document.</span></span>

### <a name="saving-your-settings"></a><span data-ttu-id="843d5-189">設定の保存</span><span class="sxs-lookup"><span data-stu-id="843d5-189">Saving your settings</span></span>

<span data-ttu-id="843d5-190">現在のセッション中に、アドインがメモリ内の設定プロパティ バッグに対して行った追加、変更、または削除を保存するには、 [Settings.saveAsync](/javascript/api/office/office.settings#saveasync-options--callback-) メソッドを呼び出してそれらの設定をドキュメントに保存する必要があります。</span><span class="sxs-lookup"><span data-stu-id="843d5-190">To save any additions, changes, or deletions your add-in made to the in-memory copy of the settings property bag during the current session, you must call the [Settings.saveAsync](/javascript/api/office/office.settings#saveasync-options--callback-) method to store them in the document.</span></span> <span data-ttu-id="843d5-191">メソッドの唯一の `saveAsync` パラメーターは _、1_ つのパラメーターを持つコールバック関数であるコールバックです。</span><span class="sxs-lookup"><span data-stu-id="843d5-191">The only parameter of the `saveAsync` method is _callback_, which is a callback function with a single parameter.</span></span>

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

<span data-ttu-id="843d5-192">コールバック パラメーターとしてメソッドに渡される匿名関数は、操作 `saveAsync` が完了すると実行されます。 </span><span class="sxs-lookup"><span data-stu-id="843d5-192">The anonymous function passed into the `saveAsync` method as the _callback_ parameter is executed when the operation is completed.</span></span> <span data-ttu-id="843d5-193">コールバック _の asyncResult_ パラメーターは、操作の状態を含 `AsyncResult` むオブジェクトへのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="843d5-193">The _asyncResult_ parameter of the callback provides access to an `AsyncResult` object that contains the status of the operation.</span></span> <span data-ttu-id="843d5-194">この例では、この関数はプロパティをチェックして、保存操作が成功したか失敗したのか確認し、その結果をアドインのページ `AsyncResult.status` に表示します。</span><span class="sxs-lookup"><span data-stu-id="843d5-194">In the example, the function checks the `AsyncResult.status` property to see if the save operation succeeded or failed, and then displays the result in the add-in's page.</span></span>

## <a name="how-to-save-custom-xml-to-the-document"></a><span data-ttu-id="843d5-195">ドキュメントにカスタム XML を保存する方法</span><span class="sxs-lookup"><span data-stu-id="843d5-195">How to save custom XML to the document</span></span>

> [!NOTE]
> <span data-ttu-id="843d5-196">このセクションでは、Word でサポートされている Office 共通 JavaScript API のコンテキストでのカスタム XML 部分について説明します。</span><span class="sxs-lookup"><span data-stu-id="843d5-196">This section discusses custom XML parts in the context of the Office Common JavaScript API which is supported in Word.</span></span> <span data-ttu-id="843d5-197">アプリケーション固有の Excel JavaScript API では、カスタム XML パーツにもアクセスできます。</span><span class="sxs-lookup"><span data-stu-id="843d5-197">The application-specific Excel JavaScript API also provides access to the custom XML parts.</span></span> <span data-ttu-id="843d5-198">Excel の API とプログラミング パターンには、わずかな違いがあります。</span><span class="sxs-lookup"><span data-stu-id="843d5-198">The Excel APIs and programming patterns are somewhat different.</span></span> <span data-ttu-id="843d5-199">詳細については、[Excel の CustomXmlPart](/javascript/api/excel/excel.customxmlpart) を参照してください。</span><span class="sxs-lookup"><span data-stu-id="843d5-199">For more information, see [Excel CustomXmlPart](/javascript/api/excel/excel.customxmlpart).</span></span>

<span data-ttu-id="843d5-200">ドキュメントの設定のサイズ制限を超える情報、または構造化文字を含む情報を格納する必要がある場合は、追加の記憶域オプションがあります。</span><span class="sxs-lookup"><span data-stu-id="843d5-200">There is an additional storage option when you need to store information that exceeds the size limits of the document Settings or which has a structured character.</span></span> <span data-ttu-id="843d5-201">Word および Excel の作業ウィンドウ アドインには、カスタムの XML マークアップを保持できます (Excel については、このセクションの冒頭にあるノートを参照してください)。</span><span class="sxs-lookup"><span data-stu-id="843d5-201">You can persist custom XML markup in a task pane add-in for Word (and for Excel, but see the note at the top of this section).</span></span> <span data-ttu-id="843d5-202">Word の場合は、[CustomXmlPart](/javascript/api/office/office.customxmlpart) とそのメソッドを使用します (繰り返しになりますが、Excel の場合は上記のノートを参照してください)。</span><span class="sxs-lookup"><span data-stu-id="843d5-202">In Word, you use the [CustomXmlPart](/javascript/api/office/office.customxmlpart) object and its methods (again, see the note above for Excel).</span></span> <span data-ttu-id="843d5-203">次のコードでは、カスタム XML パーツを作成して、その ID とコンテンツをページの div に表示します。</span><span class="sxs-lookup"><span data-stu-id="843d5-203">The following code creates a custom XML part and displays its ID and then its content in divs on the page.</span></span> <span data-ttu-id="843d5-204">XML 文字列には `xmlns` 属性が必ず存在する点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="843d5-204">Note that there must be an `xmlns` attribute in the XML string.</span></span>

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

<span data-ttu-id="843d5-p127">カスタム XML 部分を取得するには、[getByIdAsync](/javascript/api/office/office.customxmlparts#getbyidasync-id--options--callback-) メソッドを使用しますが、ID は XML 部分の作成時に生成された GUID になるため、コードの作成時に ID の内容を知ることはできません。 そのため、XML 部分を作成したら、その XML 部分の ID を設定としてすぐに保存して、覚えやすいキーを割り当てることがベスト プラクティスになります。 次のメソッドは、この方法を示してます  (ただし、カスタム設定の操作に関する詳細とベスト プラクティスについては、この記事の前半のセクションを参照してください)。</span><span class="sxs-lookup"><span data-stu-id="843d5-p127">To retrieve a custom XML part, you use the [getByIdAsync](/javascript/api/office/office.customxmlparts#getbyidasync-id--options--callback-) method, but the ID is a GUID that is generated when the XML part is created, so you can't know when coding what the ID is. For that reason, it is a good practice when creating an XML part to immediately store the ID of the XML part as a setting and give it a memorable key. The following method shows how to do this. (But see earlier sections of this article for details and best practices when working with custom settings).</span></span>

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

<span data-ttu-id="843d5-209">次のコードは、最初に設定から ID を取得することで、XML 部分を取得する方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="843d5-209">The following code shows how to retrieve the XML part by first getting its ID from a setting.</span></span>

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

## <a name="how-to-save-settings-in-an-outlook-add-in"></a><span data-ttu-id="843d5-210">Outlook アドインに設定を保存する方法</span><span class="sxs-lookup"><span data-stu-id="843d5-210">How to save settings in an Outlook add-in</span></span>

<span data-ttu-id="843d5-211">Outlook アドインに設定を保存する方法については、「Outlook アドインの状態と設定の管理」 [を参照してください](../outlook/manage-state-and-settings-outlook.md)。</span><span class="sxs-lookup"><span data-stu-id="843d5-211">For information about how to save settings in an Outlook add-in, see [Manage state and settings for an Outlook add-in](../outlook/manage-state-and-settings-outlook.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="843d5-212">関連項目</span><span class="sxs-lookup"><span data-stu-id="843d5-212">See also</span></span>

- [<span data-ttu-id="843d5-213">Office JavaScript API について</span><span class="sxs-lookup"><span data-stu-id="843d5-213">Understanding the Office JavaScript API</span></span>](understanding-the-javascript-api-for-office.md)
- [<span data-ttu-id="843d5-214">Outlook アドイン</span><span class="sxs-lookup"><span data-stu-id="843d5-214">Outlook add-ins</span></span>](../outlook/outlook-add-ins-overview.md)
- [<span data-ttu-id="843d5-215">Outlook アドインの状態と設定を管理する</span><span class="sxs-lookup"><span data-stu-id="843d5-215">Manage state and settings for an Outlook add-in</span></span>](../outlook/manage-state-and-settings-outlook.md)
- [<span data-ttu-id="843d5-216">Excel-Add-in-JavaScript-PersistCustomSettings</span><span class="sxs-lookup"><span data-stu-id="843d5-216">Excel-Add-in-JavaScript-PersistCustomSettings</span></span>](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings)
