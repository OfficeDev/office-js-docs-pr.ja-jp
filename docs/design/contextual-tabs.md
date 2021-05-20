---
title: Office アドインでカスタム コンテキスト タブを作成する
description: Office アドインにカスタム コンテキスト タブを追加する方法について説明します。
ms.date: 05/12/2021
localization_priority: Normal
ms.openlocfilehash: d03ac2c01c03353f3e2d1b54ba20616d7b42d93f
ms.sourcegitcommit: 693d364616b42eea66977eef47530adabc51a40f
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/19/2021
ms.locfileid: "52555207"
---
# <a name="create-custom-contextual-tabs-in-office-add-ins"></a><span data-ttu-id="aa046-103">Office アドインでカスタム コンテキスト タブを作成する</span><span class="sxs-lookup"><span data-stu-id="aa046-103">Create custom contextual tabs in Office Add-ins</span></span>

<span data-ttu-id="aa046-104">コンテキスト タブは、Officeのリボンの非表示のタブ コントロールで、指定したイベントがOffice ドキュメントで発生したときにタブ行に表示されます。</span><span class="sxs-lookup"><span data-stu-id="aa046-104">A contextual tab is a hidden tab control in the Office ribbon that is displayed in the tab row when a specified event occurs in the Office document.</span></span> <span data-ttu-id="aa046-105">たとえば、テーブルが選択されたときにリボンExcel表示される [テーブル **デザイン**] タブなどです。</span><span class="sxs-lookup"><span data-stu-id="aa046-105">For example, the **Table Design** tab that appears on the Excel ribbon when a table is selected.</span></span> <span data-ttu-id="aa046-106">Office アドインにカスタム コンテキスト タブを含め、表示を変更するイベント ハンドラーを作成して、表示または非表示を切り替えるタイミングを指定できます。</span><span class="sxs-lookup"><span data-stu-id="aa046-106">You can include custom contextual tabs in your Office Add-in and specify when they are visible or hidden, by creating event handlers that change the visibility.</span></span> <span data-ttu-id="aa046-107">(ただし、カスタム コンテキスト タブはフォーカスの変更に応答しません)。</span><span class="sxs-lookup"><span data-stu-id="aa046-107">(However, custom contextual tabs do not respond to focus changes.)</span></span>

> [!NOTE]
> <span data-ttu-id="aa046-108">この記事は、以下のドキュメントについて既に理解していることを前提としています。</span><span class="sxs-lookup"><span data-stu-id="aa046-108">This article assumes that you are familiar with the following documentation.</span></span> <span data-ttu-id="aa046-109">最近、アドイン コマンド (カスタム メニュー項目とリボン ボタン) を使用してない場合は、ドキュメントをご確認ください。</span><span class="sxs-lookup"><span data-stu-id="aa046-109">Please review it if you haven't worked with Add-in Commands (custom menu items and ribbon buttons) recently.</span></span>
>
> - [<span data-ttu-id="aa046-110">アドイン コマンドの基本概念</span><span class="sxs-lookup"><span data-stu-id="aa046-110">Basic concepts for Add-in Commands</span></span>](add-in-commands.md)

> [!IMPORTANT]
> <span data-ttu-id="aa046-111">カスタム コンテキスト タブは現在、Excelでのみサポートされており、これらのプラットフォームとビルドでのみサポートされています。</span><span class="sxs-lookup"><span data-stu-id="aa046-111">Custom contextual tabs are currently only supported on Excel and only on these platforms and builds:</span></span>
>
> - <span data-ttu-id="aa046-112">Windows (Microsoft 365 サブスクリプションのみ) でExcel): バージョン 2102 (ビルド 13801.20294) 以降。</span><span class="sxs-lookup"><span data-stu-id="aa046-112">Excel on Windows (Microsoft 365 subscription only): Version 2102 (Build 13801.20294) or later.</span></span>
> - <span data-ttu-id="aa046-113">Excel on the web</span><span class="sxs-lookup"><span data-stu-id="aa046-113">Excel on the web</span></span>

> [!NOTE]
> <span data-ttu-id="aa046-114">カスタム コンテキスト タブは、次の要件セットをサポートするプラットフォームでのみ機能します。</span><span class="sxs-lookup"><span data-stu-id="aa046-114">Custom contextual tabs work only on platforms that support the following requirement sets.</span></span> <span data-ttu-id="aa046-115">要件セットの詳細と、それらの要件セットの使用方法については[、「Officeアプリケーションと API 要件の指定](../develop/specify-office-hosts-and-api-requirements.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="aa046-115">For more about requirement sets and how to work with them, see [Specify Office applications and API requirements](../develop/specify-office-hosts-and-api-requirements.md).</span></span>
>
> - [<span data-ttu-id="aa046-116">リボンアピ1.2</span><span class="sxs-lookup"><span data-stu-id="aa046-116">RibbonApi 1.2</span></span>](../reference/requirement-sets/ribbon-api-requirement-sets.md)
> - [<span data-ttu-id="aa046-117">SharedRuntime 1.1</span><span class="sxs-lookup"><span data-stu-id="aa046-117">SharedRuntime 1.1</span></span>](../reference/requirement-sets/shared-runtime-requirement-sets.md)
>
> <span data-ttu-id="aa046-118">コードでランタイム チェックを使用して、ユーザーのホストとプラットフォームの組み合わせがこれらの要件[Office](../develop/specify-office-hosts-and-api-requirements.md#use-runtime-checks-in-your-javascript-code)セットをサポートしているかどうかをテストできます 。</span><span class="sxs-lookup"><span data-stu-id="aa046-118">You can use the runtime checks in your code to test whether the user's host and platform combination supports these requirement sets as described in [Specify Office applications and API requirements](../develop/specify-office-hosts-and-api-requirements.md#use-runtime-checks-in-your-javascript-code).</span></span> <span data-ttu-id="aa046-119">(マニフェストで要件セットを指定する手法は、その記事でも説明されていますが、現在のところ、RibbonApi 1.2 では機能しません)。または、 [カスタム コンテキスト タブがサポートされていない場合に、代替 UI エクスペリエンスを実装](#implement-an-alternate-ui-experience-when-custom-contextual-tabs-are-not-supported)することもできます。</span><span class="sxs-lookup"><span data-stu-id="aa046-119">(The technique of specifying the requirement sets in the manifest, which is also described in that article, does not currently work for RibbonApi 1.2.) Alternatively, you can [implement an alternate UI experience when custom contextual tabs are not supported](#implement-an-alternate-ui-experience-when-custom-contextual-tabs-are-not-supported).</span></span>

## <a name="behavior-of-custom-contextual-tabs"></a><span data-ttu-id="aa046-120">カスタム コンテキスト タブの動作</span><span class="sxs-lookup"><span data-stu-id="aa046-120">Behavior of custom contextual tabs</span></span>

<span data-ttu-id="aa046-121">カスタム コンテキスト タブのユーザー エクスペリエンスは、組み込みのOfficeコンテキスト タブのパターンに従います。</span><span class="sxs-lookup"><span data-stu-id="aa046-121">The user experience for custom contextual tabs follows the pattern of built-in Office contextual tabs.</span></span> <span data-ttu-id="aa046-122">配置カスタム コンテキスト タブの基本原則を次に示します。</span><span class="sxs-lookup"><span data-stu-id="aa046-122">The following are the basic principles for the placement custom contextual tabs:</span></span>

- <span data-ttu-id="aa046-123">カスタム コンテキスト タブが表示されている場合、リボンの右端に表示されます。</span><span class="sxs-lookup"><span data-stu-id="aa046-123">When a custom contextual tab is visible, it appears on the right end of the ribbon.</span></span>
- <span data-ttu-id="aa046-124">1 つ以上の組み込みコンテキスト タブと、アドインの 1 つ以上のカスタム コンテキスト タブが同時に表示される場合、カスタム コンテキスト タブは常にすべての組み込みコンテキスト タブの右側にあります。</span><span class="sxs-lookup"><span data-stu-id="aa046-124">If one or more built-in contextual tabs and one or more custom contextual tabs from add-ins are visible at the same time, the custom contextual tabs are always to the right of all of the built-in contextual tabs.</span></span>
- <span data-ttu-id="aa046-125">アドインに複数のコンテキスト タブがあり、複数のコンテキストが表示されている場合、アドインで定義されている順序で表示されます。</span><span class="sxs-lookup"><span data-stu-id="aa046-125">If your add-in has more than one contextual tab and there are contexts in which more than one is visible, they appear in the order in which they are defined in your add-in.</span></span> <span data-ttu-id="aa046-126">(方向はOffice言語と同じ方向、つまり、左から右に右に、右から左の言語では右から左の言語です。定義方法の詳細については、「[タブに表示されるグループとコントロール](#define-the-groups-and-controls-that-appear-on-the-tab)の定義」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="aa046-126">(The direction is the same direction as the Office language; that is, is left-to-right in left-to-right languages, but right-to-left in right-to-left languages.) See [Define the groups and controls that appear on the tab](#define-the-groups-and-controls-that-appear-on-the-tab) for details about how you define them.</span></span>
- <span data-ttu-id="aa046-127">複数のアドインに特定のコンテキストで表示されるコンテキスト タブがある場合、アドインが起動された順序で表示されます。</span><span class="sxs-lookup"><span data-stu-id="aa046-127">If more than one add-in has a contextual tab that is visible in a specific context, then they appear in the order in which the add-ins were launched.</span></span>
- <span data-ttu-id="aa046-128">カスタム *のコンテキスト* タブは、カスタム コア タブとは異なり、Office アプリケーションのリボンに永続的に追加されません。</span><span class="sxs-lookup"><span data-stu-id="aa046-128">Custom *contextual* tabs, unlike custom core tabs, are not added permanently to the Office application's ribbon.</span></span> <span data-ttu-id="aa046-129">これらのファイルは、アドインが実行されているドキュメントOfficeにのみ存在します。</span><span class="sxs-lookup"><span data-stu-id="aa046-129">They are present only in Office documents on which your add-in is running.</span></span>

## <a name="major-steps-for-including-a-contextual-tab-in-an-add-in"></a><span data-ttu-id="aa046-130">アドインにコンテキスト タブを含める主な手順</span><span class="sxs-lookup"><span data-stu-id="aa046-130">Major steps for including a contextual tab in an add-in</span></span>

<span data-ttu-id="aa046-131">アドインにカスタム コンテキスト タブを含める主な手順を次に示します。</span><span class="sxs-lookup"><span data-stu-id="aa046-131">The following are the major steps for including a custom contextual tab in an add-in:</span></span>

1. <span data-ttu-id="aa046-132">共有ランタイムを使用するようにアドインを構成します。</span><span class="sxs-lookup"><span data-stu-id="aa046-132">Configure the add-in to use a shared runtime.</span></span>
1. <span data-ttu-id="aa046-133">タブ、およびタブに表示されるグループとコントロールを定義します。</span><span class="sxs-lookup"><span data-stu-id="aa046-133">Define the tab and the groups and controls that appear on it.</span></span>
1. <span data-ttu-id="aa046-134">コンテキスト タブをOfficeに登録します。</span><span class="sxs-lookup"><span data-stu-id="aa046-134">Register the contextual tab with Office.</span></span>
1. <span data-ttu-id="aa046-135">タブが表示される状況を指定します。</span><span class="sxs-lookup"><span data-stu-id="aa046-135">Specify the circumstances when the tab will be visible.</span></span>

## <a name="configure-the-add-in-to-use-a-shared-runtime"></a><span data-ttu-id="aa046-136">共有ランタイムを使用するようにアドインを構成する</span><span class="sxs-lookup"><span data-stu-id="aa046-136">Configure the add-in to use a shared runtime</span></span>

<span data-ttu-id="aa046-137">カスタム コンテキスト タブを追加するには、共有ランタイムを使用するアドインが必要です。</span><span class="sxs-lookup"><span data-stu-id="aa046-137">Adding custom contextual tabs requires your add-in to use the shared runtime.</span></span> <span data-ttu-id="aa046-138">詳細については、「 [共有ランタイムを使用するようにアドインを構成する](../develop/configure-your-add-in-to-use-a-shared-runtime.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="aa046-138">For more information, see [Configure an add-in to use a shared runtime](../develop/configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

## <a name="define-the-groups-and-controls-that-appear-on-the-tab"></a><span data-ttu-id="aa046-139">タブに表示されるグループとコントロールを定義する</span><span class="sxs-lookup"><span data-stu-id="aa046-139">Define the groups and controls that appear on the tab</span></span>

<span data-ttu-id="aa046-140">マニフェストで XML で定義されるカスタム コア タブとは異なり、カスタム コンテキスト タブは JSON BLOB を使用して実行時に定義されます。</span><span class="sxs-lookup"><span data-stu-id="aa046-140">Unlike custom core tabs, which are defined with XML in the manifest, custom contextual tabs are defined at runtime with a JSON blob.</span></span> <span data-ttu-id="aa046-141">コードは、Blob を JavaScript オブジェクトに解析し、そのオブジェクトを[Office.ribbon.requestCreateControls](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls-tabDefinition-)メソッドに渡します。</span><span class="sxs-lookup"><span data-stu-id="aa046-141">Your code parses the blob into a JavaScript object, and then passes the object to the [Office.ribbon.requestCreateControls](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls-tabDefinition-) method.</span></span> <span data-ttu-id="aa046-142">カスタム コンテキスト タブは、アドインが現在実行されているドキュメントにのみ表示されます。</span><span class="sxs-lookup"><span data-stu-id="aa046-142">Custom contextual tabs are only present in documents on which your add-in is currently running.</span></span> <span data-ttu-id="aa046-143">これは、アドインがインストールされるときにOffice アプリケーション リボンに追加され、別のドキュメントが開かれたときに表示されたままになるカスタム コア タブとは異なります。</span><span class="sxs-lookup"><span data-stu-id="aa046-143">This is different from custom core tabs which are added to the Office application ribbon when the add-in is installed and remain present when another document is opened.</span></span> <span data-ttu-id="aa046-144">また、 `requestCreateControls` このメソッドは、アドインのセッションで 1 回だけ実行できます。</span><span class="sxs-lookup"><span data-stu-id="aa046-144">Also, the `requestCreateControls` method can be run only once in a session of your add-in.</span></span> <span data-ttu-id="aa046-145">再度呼び出されると、エラーがスローされます。</span><span class="sxs-lookup"><span data-stu-id="aa046-145">If it is called again, an error is thrown.</span></span>

> [!NOTE]
> <span data-ttu-id="aa046-146">JSON BLOB のプロパティとサブプロパティ (およびキー名) の構造は、マニフェスト XML の [CustomTab](../reference/manifest/customtab.md) 要素とその子孫要素の構造とほぼ平行です。</span><span class="sxs-lookup"><span data-stu-id="aa046-146">The structure of the JSON blob's properties and subproperties (and the key names) is roughly parallel to the structure of the [CustomTab](../reference/manifest/customtab.md) element and its descendant elements in the manifest XML.</span></span>

<span data-ttu-id="aa046-147">コンテキスト タブ JSON BLOB の例を段階的に作成します。</span><span class="sxs-lookup"><span data-stu-id="aa046-147">We'll construct an example of a contextual tabs JSON blob step-by-step.</span></span> <span data-ttu-id="aa046-148">コンテキスト タブ JSON の完全なスキーマは、 [dynamic-ribbon.schema.jsにあります](https://developer.microsoft.com/json-schemas/office-js/dynamic-ribbon.schema.json)。</span><span class="sxs-lookup"><span data-stu-id="aa046-148">The full schema for the contextual tab JSON is at [dynamic-ribbon.schema.json](https://developer.microsoft.com/json-schemas/office-js/dynamic-ribbon.schema.json).</span></span> <span data-ttu-id="aa046-149">Visual Studio Codeで作業している場合は、このファイルを使用してIntelliSenseを取得し、JSON を検証できます。</span><span class="sxs-lookup"><span data-stu-id="aa046-149">If you are working in Visual Studio Code, you can use this file to get IntelliSense and to validate your JSON.</span></span> <span data-ttu-id="aa046-150">詳細については[、「json の編集 」を参照してくださいVisual Studio Code - JSON スキーマと設定を](https://code.visualstudio.com/docs/languages/json#_json-schemas-and-settings)使用します。</span><span class="sxs-lookup"><span data-stu-id="aa046-150">For more information, see [Editing JSON with Visual Studio Code - JSON schemas and settings](https://code.visualstudio.com/docs/languages/json#_json-schemas-and-settings).</span></span>


1. <span data-ttu-id="aa046-151">まず、 という名前の 2 つの配列プロパティを持つ JSON 文字列 `actions` を作成 `tabs` します。</span><span class="sxs-lookup"><span data-stu-id="aa046-151">Begin by creating a JSON string with two array properties named `actions` and `tabs`.</span></span> <span data-ttu-id="aa046-152">`actions`配列は、コンテキスト タブのコントロールで実行できるすべての関数の仕様です。`tabs`配列は、*最大 20* 個までの 1 つ以上のコンテキスト タブを定義します。</span><span class="sxs-lookup"><span data-stu-id="aa046-152">The `actions` array is a specification of all the functions that can be executed by controls on the contextual tab. The `tabs` array defines one or more contextual tabs, *up to a maximum of 20*.</span></span>

    ```json
    '{
      "actions": [

      ],
      "tabs": [

      ]
    }'
    ```

1. <span data-ttu-id="aa046-153">この単純なコンテキスト タブの例では、ボタンは 1 つだけで、単一のアクションのみが表示されます。</span><span class="sxs-lookup"><span data-stu-id="aa046-153">This simple example of a contextual tab will have only a single button and, thus, only a single action.</span></span> <span data-ttu-id="aa046-154">次の要素を配列の唯一のメンバーとして追加 `actions` します。</span><span class="sxs-lookup"><span data-stu-id="aa046-154">Add the following as the only member of the `actions` array.</span></span> <span data-ttu-id="aa046-155">このマークアップについて、次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="aa046-155">About this markup, note:</span></span>

    - <span data-ttu-id="aa046-156">`id`および `type` プロパティは必須です。</span><span class="sxs-lookup"><span data-stu-id="aa046-156">The `id` and `type` properties are mandatory.</span></span>
    - <span data-ttu-id="aa046-157">の値 `type` は、"関数の実行" または "タスク ウィンドウの表示" のいずれかです。</span><span class="sxs-lookup"><span data-stu-id="aa046-157">The value of `type` can be either "ExecuteFunction" or "ShowTaskpane".</span></span>
    - <span data-ttu-id="aa046-158">`functionName`プロパティは、 の値が の場合にのみ使用 `type` されます `ExecuteFunction` 。</span><span class="sxs-lookup"><span data-stu-id="aa046-158">The `functionName` property is only used when the value of `type` is `ExecuteFunction`.</span></span> <span data-ttu-id="aa046-159">これは、関数ファイルで定義されている関数の名前です。</span><span class="sxs-lookup"><span data-stu-id="aa046-159">It is the name of a function defined in the FunctionFile.</span></span> <span data-ttu-id="aa046-160">FunctionFile の詳細については、「アドイン [コマンドの基本概念](add-in-commands.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="aa046-160">For more information about the FunctionFile, see [Basic concepts for Add-in Commands](add-in-commands.md).</span></span>
    - <span data-ttu-id="aa046-161">後の手順で、このアクションをコンテキスト タブのボタンにマップします。</span><span class="sxs-lookup"><span data-stu-id="aa046-161">In a later step, you will map this action to a button on the contextual tab.</span></span>

    ```json
    {
      "id": "executeWriteData",
      "type": "ExecuteFunction",
      "functionName": "writeData"
    }
   ```

1. <span data-ttu-id="aa046-162">次の要素を配列の唯一のメンバーとして追加 `tabs` します。</span><span class="sxs-lookup"><span data-stu-id="aa046-162">Add the following as the only member of the `tabs` array.</span></span> <span data-ttu-id="aa046-163">このマークアップについて、次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="aa046-163">About this markup, note:</span></span>

    - <span data-ttu-id="aa046-164">`id` プロパティは必須です。</span><span class="sxs-lookup"><span data-stu-id="aa046-164">The `id` property is required.</span></span> <span data-ttu-id="aa046-165">アドイン内のすべてのコンテキスト タブに固有の簡単な説明 ID を使用します。</span><span class="sxs-lookup"><span data-stu-id="aa046-165">Use a brief, descriptive ID that is unique among all contextual tabs in your add-in.</span></span>
    - <span data-ttu-id="aa046-166">`label` プロパティは必須です。</span><span class="sxs-lookup"><span data-stu-id="aa046-166">The `label` property is required.</span></span> <span data-ttu-id="aa046-167">コンテキスト タブのラベルとして使用するわかりやすい文字列です。</span><span class="sxs-lookup"><span data-stu-id="aa046-167">It is a user-friendly string to serve as the label of the contextual tab.</span></span>
    - <span data-ttu-id="aa046-168">`groups` プロパティは必須です。</span><span class="sxs-lookup"><span data-stu-id="aa046-168">The `groups` property is required.</span></span> <span data-ttu-id="aa046-169">タブに表示されるコントロールのグループを定義します。少なくとも 1 つのメンバー *と 20 以下の メンバーが* 必要です。</span><span class="sxs-lookup"><span data-stu-id="aa046-169">It defines the groups of controls that will appear on the tab. It must have at least one member *and no more than 20*.</span></span> <span data-ttu-id="aa046-170">(カスタム コンテキスト タブに設定できるコントロールの数にも制限があり、グループの数も制限されます。</span><span class="sxs-lookup"><span data-stu-id="aa046-170">(There are also limits on the number of controls that you can have on a custom contextual tab and that will also constrain how many groups that you have.</span></span> <span data-ttu-id="aa046-171">詳細については、次の手順を参照してください。</span><span class="sxs-lookup"><span data-stu-id="aa046-171">See the next step for more information.)</span></span>

    > [!NOTE]
    > <span data-ttu-id="aa046-172">Tab オブジェクトには、 `visible` アドインの起動時にタブをすぐに表示するかどうかを指定するオプションのプロパティを持つ場合もあります。</span><span class="sxs-lookup"><span data-stu-id="aa046-172">The tab object can also have an optional `visible` property that specifies whether the tab is visible immediately when the add-in starts up.</span></span> <span data-ttu-id="aa046-173">コンテキスト タブは通常、ユーザー イベントが表示をトリガーするまで非表示になるため (ユーザーがドキュメント内の何らかの種類のエンティティを選択した場合など)、 `visible` プロパティは既定で `false` 表示されない場合に設定されます。</span><span class="sxs-lookup"><span data-stu-id="aa046-173">Since contextual tabs are normally hidden until a user event triggers their visibility (such as the user selecting an entity of some type in the document), the `visible` property defaults to `false` when not present.</span></span> <span data-ttu-id="aa046-174">後のセクションでは、イベントに応答してプロパティを設定 `true` する方法を示します。</span><span class="sxs-lookup"><span data-stu-id="aa046-174">In a later section, we show how to set the property to `true` in response to an event.</span></span>

    ```json
    {
      "id": "CtxTab1",
      "label": "Contoso Data",
      "groups": [

      ]
    }
    ```

1. <span data-ttu-id="aa046-175">簡単な進行中の例では、コンテキスト タブには 1 つのグループしかありません。</span><span class="sxs-lookup"><span data-stu-id="aa046-175">In the simple ongoing example, the contextual tab has only a single group.</span></span> <span data-ttu-id="aa046-176">次の要素を配列の唯一のメンバーとして追加 `groups` します。</span><span class="sxs-lookup"><span data-stu-id="aa046-176">Add the following as the only member of the `groups` array.</span></span> <span data-ttu-id="aa046-177">このマークアップについて、次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="aa046-177">About this markup, note:</span></span>

    - <span data-ttu-id="aa046-178">すべてのプロパティが必要です。</span><span class="sxs-lookup"><span data-stu-id="aa046-178">All the properties are required.</span></span>
    - <span data-ttu-id="aa046-179">`id`プロパティは、タブ内のすべてのグループ間で一意である必要があります。</span><span class="sxs-lookup"><span data-stu-id="aa046-179">The `id` property must be unique among all the groups in the tab. Use a brief, descriptive ID.</span></span>
    - <span data-ttu-id="aa046-180">`label`は、グループのラベルとして使用するわかりやすい文字列です。</span><span class="sxs-lookup"><span data-stu-id="aa046-180">The `label` is a user-friendly string to serve as the label of the group.</span></span>
    - <span data-ttu-id="aa046-181">`icon`プロパティの値は、リボンのサイズとアプリケーション ウィンドウに応じてリボンにグループが表示されるアイコンを指定するオブジェクトの配列Office。</span><span class="sxs-lookup"><span data-stu-id="aa046-181">The `icon` property's value is an array of objects that specify the icons that the group will have on the ribbon depending on the size of the ribbon and the Office application window.</span></span>
    - <span data-ttu-id="aa046-182">`controls`プロパティの値は、グループ内のボタンとメニューを指定するオブジェクトの配列です。</span><span class="sxs-lookup"><span data-stu-id="aa046-182">The `controls` property's value is an array of objects that specify the buttons and menus in the group.</span></span> <span data-ttu-id="aa046-183">少なくとも 1 つ存在する必要があります。</span><span class="sxs-lookup"><span data-stu-id="aa046-183">There must be at least one.</span></span>

    > [!IMPORTANT]
    > <span data-ttu-id="aa046-184">*タブ全体のコントロールの合計数は 20 個以下にできます。*</span><span class="sxs-lookup"><span data-stu-id="aa046-184">*The total number of controls on the whole tab can be no more than 20.*</span></span> <span data-ttu-id="aa046-185">たとえば、それぞれ 6 つのコントロールを持つ 3 つのグループと、4 番目のグループに 2 つのコントロールを含め、それぞれ 6 つのコントロールを持つ 4 つのグループを持つことはできません。</span><span class="sxs-lookup"><span data-stu-id="aa046-185">For example, you could have 3 groups with 6 controls each, and a fourth group with 2 controls, but you cannot have 4 groups with 6 controls each.</span></span>  

    ```json
    {
        "id": "CustomGroup111",
        "label": "Insertion",
        "icon": [

        ],
        "controls": [

        ]
    }
    ```

1. <span data-ttu-id="aa046-186">各グループには、32x32 pxと80x80 pxの2つ以上のサイズのアイコンが必要です。</span><span class="sxs-lookup"><span data-stu-id="aa046-186">Every group must have an icon of at least two sizes, 32x32 px and 80x80 px.</span></span> <span data-ttu-id="aa046-187">オプションで、サイズ 16x16 px、20x20 px、24x24 px、40x40 px、48x48 px、および 64x64 ピクセルのアイコンを使用することもできます。</span><span class="sxs-lookup"><span data-stu-id="aa046-187">Optionally, you can also have icons of sizes 16x16 px, 20x20 px, 24x24 px, 40x40 px, 48x48 px, and 64x64 px.</span></span> <span data-ttu-id="aa046-188">Office、リボンのサイズとアプリケーション ウィンドウに基づいて、使用するアイコンOffice決定します。</span><span class="sxs-lookup"><span data-stu-id="aa046-188">Office decides which icon to use based on the size of the ribbon and Office application window.</span></span> <span data-ttu-id="aa046-189">アイコン配列に次のオブジェクトを追加します。</span><span class="sxs-lookup"><span data-stu-id="aa046-189">Add the following objects to the icon array.</span></span> <span data-ttu-id="aa046-190">(ウィンドウとリボンのサイズが *グループのコントロール* の少なくとも 1 つが表示されるのに十分な大きさの場合は、グループ アイコンはまったく表示されません。</span><span class="sxs-lookup"><span data-stu-id="aa046-190">(If the window and ribbon sizes are large enough for at least one of the *controls* on the group to appear, then no group icon at all appears.</span></span> <span data-ttu-id="aa046-191">たとえば、Word ウィンドウを縮小および展開するときに、Word リボンの **[スタイル]** グループを確認します。このマークアップについて、次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="aa046-191">For an example, watch the **Styles** group on the Word ribbon as you shrink and expand the Word window.) About this markup, note:</span></span>

    - <span data-ttu-id="aa046-192">両方のプロパティが必要です。</span><span class="sxs-lookup"><span data-stu-id="aa046-192">Both the properties are required.</span></span>
    - <span data-ttu-id="aa046-193">`size`プロパティの単位はピクセルです。</span><span class="sxs-lookup"><span data-stu-id="aa046-193">The `size` property unit of measure is pixels.</span></span> <span data-ttu-id="aa046-194">アイコンは常に正方形であるため、数値は高さと幅の両方になります。</span><span class="sxs-lookup"><span data-stu-id="aa046-194">Icons are always square, so the number is both the height and the width.</span></span>
    - <span data-ttu-id="aa046-195">`sourceLocation`プロパティは、アイコンへの完全な URL を指定します。</span><span class="sxs-lookup"><span data-stu-id="aa046-195">The `sourceLocation` property specifies the full URL to the icon.</span></span>

    > [!IMPORTANT]
    > <span data-ttu-id="aa046-196">開発から運用環境に移行する場合 (localhost から contoso.com にドメインを変更するなど)、アドインのマニフェストの URL を通常変更する必要があるのと同様に、コンテキスト タブ JSON の URL も変更する必要があります。</span><span class="sxs-lookup"><span data-stu-id="aa046-196">Just as you typically must change the URLs in the add-in's manifest when you move from development to production (such as changing the domain from localhost to contoso.com), you must also change the URLs in your contextual tabs JSON.</span></span>

    ```json
    {
        "size": 32,
        "sourceLocation": "https://cdn.contoso.com/addins/datainsertion/Images/Group32x32.png"
    },
    {
        "size": 80,
        "sourceLocation": "https://cdn.contoso.com/addins/datainsertion/Images/Group80x80.png"
    }
    ```

1. <span data-ttu-id="aa046-197">簡単な進行中の例では、グループにはボタンが 1 つしかありません。</span><span class="sxs-lookup"><span data-stu-id="aa046-197">In our simple ongoing example, the group has only a single button.</span></span> <span data-ttu-id="aa046-198">次のオブジェクトを配列の唯一のメンバーとして追加 `controls` します。</span><span class="sxs-lookup"><span data-stu-id="aa046-198">Add the following object as the only member of the `controls` array.</span></span> <span data-ttu-id="aa046-199">このマークアップについて、次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="aa046-199">About this markup, note:</span></span>

    - <span data-ttu-id="aa046-200">を除くすべてのプロパティ `enabled` が必要です。</span><span class="sxs-lookup"><span data-stu-id="aa046-200">All the properties, except `enabled`, are required.</span></span>
    - <span data-ttu-id="aa046-201">`type` コントロールの種類を指定します。</span><span class="sxs-lookup"><span data-stu-id="aa046-201">`type` specifies the type of control.</span></span> <span data-ttu-id="aa046-202">値は"ボタン"、"メニュー"、または"モバイルボタン"にすることができます。</span><span class="sxs-lookup"><span data-stu-id="aa046-202">The values can be "Button", "Menu", or "MobileButton".</span></span>
    - <span data-ttu-id="aa046-203">`id` 最大 125 文字まで指定できます。</span><span class="sxs-lookup"><span data-stu-id="aa046-203">`id` can be up to 125 characters.</span></span> 
    - <span data-ttu-id="aa046-204">`actionId` は、配列で定義されたアクションの ID でなければなりません `actions` 。</span><span class="sxs-lookup"><span data-stu-id="aa046-204">`actionId` must be the ID of an action defined in the `actions` array.</span></span> <span data-ttu-id="aa046-205">(このセクションのステップ 1 を参照してください。</span><span class="sxs-lookup"><span data-stu-id="aa046-205">(See step 1 of this section.)</span></span>
    - <span data-ttu-id="aa046-206">`label` は、ボタンのラベルとして使用するわかりやすい文字列です。</span><span class="sxs-lookup"><span data-stu-id="aa046-206">`label` is a user-friendly string to serve as the label of the button.</span></span>
    - <span data-ttu-id="aa046-207">`superTip` は、ツール ヒントの豊富な形式を表します。</span><span class="sxs-lookup"><span data-stu-id="aa046-207">`superTip` represents a rich form of tool tip.</span></span> <span data-ttu-id="aa046-208">`title`プロパティと プロパティ `description` の両方が必要です。</span><span class="sxs-lookup"><span data-stu-id="aa046-208">Both the `title` and `description` properties are required.</span></span>
    - <span data-ttu-id="aa046-209">`icon` ボタンのアイコンを指定します。</span><span class="sxs-lookup"><span data-stu-id="aa046-209">`icon` specifies the icons for the button.</span></span> <span data-ttu-id="aa046-210">グループアイコンに関する以前の解説もここに当てはまります。</span><span class="sxs-lookup"><span data-stu-id="aa046-210">The previous remarks about the group icon apply here too.</span></span>
    - <span data-ttu-id="aa046-211">`enabled` (オプション)は、コンテキストタブが表示されたときにボタンを有効にするかどうかを指定します。</span><span class="sxs-lookup"><span data-stu-id="aa046-211">`enabled` (optional) specifies whether the button is enabled when the contextual tab appears starts up.</span></span> <span data-ttu-id="aa046-212">存在しない場合のデフォルトは `true` です。</span><span class="sxs-lookup"><span data-stu-id="aa046-212">The default if not present is `true`.</span></span> 

    ```json
    {
        "type": "Button",
        "id": "CtxBt112",
        "actionId": "executeWriteData",
        "enabled": false,
        "label": "Write Data",
        "superTip": {
            "title": "Data Insertion",
            "description": "Use this button to insert data into the document."
        },
        "icon": [
            {
                "size": 32,
                "sourceLocation": "https://cdn.contoso.com/addins/datainsertion/Images/WriteDataButton32x32.png"
            },
            {
                "size": 80,
                "sourceLocation": "https://cdn.contoso.com/addins/datainsertion/Images/WriteDataButton80x80.png"
            }
        ]
    }
    ```
 
<span data-ttu-id="aa046-213">JSON BLOB の完全な例を次に示します。</span><span class="sxs-lookup"><span data-stu-id="aa046-213">The following is the complete example of the JSON blob:</span></span>

```json
`{
  "actions": [
    {
      "id": "executeWriteData",
      "type": "ExecuteFunction",
      "functionName": "writeData"
    }
  ],
  "tabs": [
    {
      "id": "CtxTab1",
      "label": "Contoso Data",
      "groups": [
        {
          "id": "CustomGroup111",
          "label": "Insertion",
          "icon": [
            {
                "size": 32,
                "sourceLocation": "https://cdn.contoso.com/addins/datainsertion/Images/Group32x32.png"
            },
            {
                "size": 80,
                "sourceLocation": "https://cdn.contoso.com/addins/datainsertion/Images/Group80x80.png"
            }
          ],
          "controls": [
            {
                "type": "Button",
                "id": "CtxBt112",
                "actionId": "executeWriteData",
                "enabled": false,
                "label": "Write Data",
                "superTip": {
                    "title": "Data Insertion",
                    "description": "Use this button to insert data into the document."
                },
                "icon": [
                    {
                        "size": 32,
                        "sourceLocation": "https://cdn.contoso.com/addins/datainsertion/Images/WriteDataButton32x32.png"
                    },
                    {
                        "size": 80,
                        "sourceLocation": "https://cdn.contoso.com/addins/datainsertion/Images/WriteDataButton80x80.png"
                    }
                ]
            }
          ]
        }
      ]
    }
  ]
}`
```

## <a name="register-the-contextual-tab-with-office-with-requestcreatecontrols"></a><span data-ttu-id="aa046-214">コンテキスト タブをOfficeに登録します。</span><span class="sxs-lookup"><span data-stu-id="aa046-214">Register the contextual tab with Office with requestCreateControls</span></span>

<span data-ttu-id="aa046-215">コンテキスト タブは[、Office.ribbon.requestCreateControls](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls_tabDefinition_)メソッドを呼び出すことによってOfficeに登録されます。</span><span class="sxs-lookup"><span data-stu-id="aa046-215">The contextual tab is registered with Office by calling the [Office.ribbon.requestCreateControls](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls_tabDefinition_) method.</span></span> <span data-ttu-id="aa046-216">これは通常、メソッドに割り当てられている関数 `Office.initialize` またはメソッドを使用して実行されます `Office.onReady` 。</span><span class="sxs-lookup"><span data-stu-id="aa046-216">This is typically done in either the function that is assigned to `Office.initialize` or with the `Office.onReady` method.</span></span> <span data-ttu-id="aa046-217">これらのメソッドとアドインの初期化の詳細については[、「Office アドインの初期化」を](../develop/initialize-add-in.md)参照してください。</span><span class="sxs-lookup"><span data-stu-id="aa046-217">For more about these methods and initializing the add-in, see [Initialize your Office Add-in](../develop/initialize-add-in.md).</span></span> <span data-ttu-id="aa046-218">ただし、初期化後はいつでもメソッドを呼び出すことができます。</span><span class="sxs-lookup"><span data-stu-id="aa046-218">You can, however, call the method anytime after initialization.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="aa046-219">`requestCreateControls`このメソッドは、アドインの特定のセッションで 1 回だけ呼び出すことができます。</span><span class="sxs-lookup"><span data-stu-id="aa046-219">The `requestCreateControls` method can be called only once in a given session of an add-in.</span></span> <span data-ttu-id="aa046-220">再び呼び出されると、エラーがスローされます。</span><span class="sxs-lookup"><span data-stu-id="aa046-220">An error is thrown if it is called again.</span></span>

<span data-ttu-id="aa046-221">次に例を示します。</span><span class="sxs-lookup"><span data-stu-id="aa046-221">The following is an example.</span></span> <span data-ttu-id="aa046-222">JSON 文字列を JavaScript 関数に渡す前に、メソッドを使用して `JSON.parse` JavaScript オブジェクトに変換する必要があることに注意してください。</span><span class="sxs-lookup"><span data-stu-id="aa046-222">Note that the JSON string must be converted to a JavaScript object with the `JSON.parse` method before it can be passed to a JavaScript function.</span></span>

```javascript
Office.onReady(async () => {
    const contextualTabJSON = ` ... `; // Assign the JSON string such as the one at the end of the preceding section.
    const contextualTab = JSON.parse(contextualTabJSON);
    await Office.ribbon.requestCreateControls(contextualTab);
});
```

## <a name="specify-the-contexts-when-the-tab-will-be-visible-with-requestupdate"></a><span data-ttu-id="aa046-223">タブが requestUpdate で表示されるコンテキストを指定します。</span><span class="sxs-lookup"><span data-stu-id="aa046-223">Specify the contexts when the tab will be visible with requestUpdate</span></span>

<span data-ttu-id="aa046-224">通常、ユーザーが開始したイベントがアドインコンテキストを変更したときに、カスタム コンテキスト タブが表示されます。</span><span class="sxs-lookup"><span data-stu-id="aa046-224">Typically, a custom contextual tab should appear when a user-initiated event changes the add-in context.</span></span> <span data-ttu-id="aa046-225">グラフ (Excel ブックの既定のワークシート) がアクティブになったときに、タブを表示する必要があるシナリオを考えてみます。</span><span class="sxs-lookup"><span data-stu-id="aa046-225">Consider a scenario in which the tab should be visible when, and only when, a chart (on the default worksheet of an Excel workbook) is activated.</span></span>

<span data-ttu-id="aa046-226">まず、ハンドラを割り当てます。</span><span class="sxs-lookup"><span data-stu-id="aa046-226">Begin by assigning handlers.</span></span> <span data-ttu-id="aa046-227">これは、 `Office.onReady` 通常、このメソッドで、後の手順で作成したハンドラーを、 ワークシート `onActivated` 内のすべてのグラフの イベントと に割り当てる方法 `onDeactivated` と同様に行われます。</span><span class="sxs-lookup"><span data-stu-id="aa046-227">This is commonly done in the `Office.onReady` method as in the following example which assigns handlers (created in a later step) to the `onActivated` and `onDeactivated` events of all the charts in the worksheet.</span></span>

```javascript
Office.onReady(async () => {
    const contextualTabJSON = ` ... `; // Assign the JSON string.
    const contextualTab = JSON.parse(contextualTabJSON);
    await Office.ribbon.requestCreateControls(contextualTab);

    await Excel.run(context => {
        var charts = context.workbook.worksheets
            .getActiveWorksheet()
            .charts;
        charts.onActivated.add(showDataTab);
        charts.onDeactivated.add(hideDataTab);
        return context.sync();
    });
});
```

<span data-ttu-id="aa046-228">次に、ハンドラーを定義します。</span><span class="sxs-lookup"><span data-stu-id="aa046-228">Next, define the handlers.</span></span> <span data-ttu-id="aa046-229">次に示す簡単な例を `showDataTab` 示しますが、関数のより堅牢なバージョンについては、この記事の後の [「HostRestartNeeded エラーの処理](#handle-the-hostrestartneeded-error) 」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="aa046-229">The following is a simple example of a `showDataTab`, but see [Handling the HostRestartNeeded error](#handle-the-hostrestartneeded-error) later in this article for a more robust version of the function.</span></span> <span data-ttu-id="aa046-230">このコードについては、以下の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="aa046-230">About this code, note:</span></span>

- <span data-ttu-id="aa046-231">Office では、リボンの状態を更新するタイミングが制御されます。</span><span class="sxs-lookup"><span data-stu-id="aa046-231">Office controls when it updates the state of the ribbon.</span></span> <span data-ttu-id="aa046-232">[Office.ribbon.requestUpdate](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestupdate-input-)メソッドは、更新要求をキューに入れます。</span><span class="sxs-lookup"><span data-stu-id="aa046-232">The  [Office.ribbon.requestUpdate](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestupdate-input-) method queues a request to update.</span></span> <span data-ttu-id="aa046-233">このメソッドは、 `Promise` リボンが実際に更新されたときではなく、要求をキューに入れるとすぐにオブジェクトを解決します。</span><span class="sxs-lookup"><span data-stu-id="aa046-233">The method will resolve the `Promise` object as soon as it has queued the request, not when the ribbon actually updates.</span></span>
- <span data-ttu-id="aa046-234">メソッドのパラメーター `requestUpdate` は、(1) *JSON で指定されたとおり* にタブを ID で指定し、(2) タブの可視性を指定する [RibbonUpdaterData](/javascript/api/office/office.ribbonupdaterdata)オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="aa046-234">The parameter for the `requestUpdate` method is a [RibbonUpdaterData](/javascript/api/office/office.ribbonupdaterdata) object that (1) specifies the tab by its ID *exactly as specified in the JSON* and (2) specifies visibility of the tab.</span></span>
- <span data-ttu-id="aa046-235">同じコンテキストで表示する必要があるカスタム コンテキスト タブが複数ある場合は、単に配列にタブ オブジェクトを追加するだけです `tabs` 。</span><span class="sxs-lookup"><span data-stu-id="aa046-235">If you have more than one custom contextual tab that should be visible in the same context, you simply add additional tab objects to the `tabs` array.</span></span>

```javascript
async function showDataTab() {
    await Office.ribbon.requestUpdate({
        tabs: [
            {
                id: "CtxTab1",
                visible: true
            }
        ]});
}
```

<span data-ttu-id="aa046-236">タブを非表示にするハンドラーはほぼ同じですが、プロパティを `visible` に戻す点が異なっています `false` 。</span><span class="sxs-lookup"><span data-stu-id="aa046-236">The handler to hide the tab is nearly identical, except that it sets the `visible` property back to `false`.</span></span>

<span data-ttu-id="aa046-237">JavaScript ライブラリOfficeには、オブジェクトの構築を容易にするためにいくつかのインターフェイス (型) も用意 `RibbonUpdateData` されています。</span><span class="sxs-lookup"><span data-stu-id="aa046-237">The Office JavaScript library also provides several interfaces (types) to make it easier to construct the`RibbonUpdateData` object.</span></span> <span data-ttu-id="aa046-238">TypeScript の `showDataTab` 関数を次に示します。</span><span class="sxs-lookup"><span data-stu-id="aa046-238">The following is the `showDataTab` function in TypeScript and it makes use of these types.</span></span>

```typescript
const showDataTab = async () => {
    const myContextualTab: Office.Tab = {id: "CtxTab1", visible: true};
    const ribbonUpdater: Office.RibbonUpdaterData = { tabs: [ myContextualTab ]};
    await Office.ribbon.requestUpdate(ribbonUpdater);
}
```

### <a name="toggle-tab-visibility-and-the-enabled-status-of-a-button-at-the-same-time"></a><span data-ttu-id="aa046-239">タブの表示とボタンの有効ステータスを同時に切り替える</span><span class="sxs-lookup"><span data-stu-id="aa046-239">Toggle tab visibility and the enabled status of a button at the same time</span></span>

<span data-ttu-id="aa046-240">この `requestUpdate` メソッドは、カスタム コンテキスト タブまたはカスタム コア タブでカスタム ボタンの有効または無効の状態を切り替える場合にも使用されます。詳細については、「 アドイン [コマンドの有効化と無効化](disable-add-in-commands.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="aa046-240">The `requestUpdate` method is also used to toggle the enabled or disabled status of a custom button on either a custom contextual tab or a custom core tab. For details about this, see [Enable and Disable Add-in Commands](disable-add-in-commands.md).</span></span> <span data-ttu-id="aa046-241">タブの表示とボタンの有効な状態の両方を同時に変更するシナリオが考えられる場合があります。</span><span class="sxs-lookup"><span data-stu-id="aa046-241">There may be scenarios in which you want to change both the visibility of a tab and the enabled status of a button at the same time.</span></span> <span data-ttu-id="aa046-242">これは、 の 1 回の呼び出しで行うことができます `requestUpdate` 。</span><span class="sxs-lookup"><span data-stu-id="aa046-242">You can do this with a single call of `requestUpdate`.</span></span> <span data-ttu-id="aa046-243">次の例では、コンテキスト タブが表示されるようにすると同時に、コア タブのボタンが有効になります。</span><span class="sxs-lookup"><span data-stu-id="aa046-243">The following is an example in which a button on a core tab is enabled at the same time as a contextual tab is made visible.</span></span>

```javascript
function myContextChanges() {
    Office.ribbon.requestUpdate({
        tabs: [
            {
                id: "CtxTab1",
                visible: true
            },
            {
                id: "OfficeAppTab1",
                groups: [
                    {
                        id: "CustomGroup111",
                        controls: [
                            {
                                id: "MyButton",
                                enabled: true
                            }
                        ]
                    }
                ]
            ]}
        ]
    });
}
```

<span data-ttu-id="aa046-244">次の例では、有効になっているボタンは、表示されているコンテキスト タブとまったく同じです。</span><span class="sxs-lookup"><span data-stu-id="aa046-244">In the following example, the button that is enabled is on the very same contextual tab that is being made visible.</span></span>

```javascript
function myContextChanges() {
    Office.ribbon.requestUpdate({
        tabs: [
            {
                id: "CtxTab1",
                visible: true,
                groups: [
                    {
                        id: "CustomGroup111",
                        controls: [
                            {
                                id: "MyButton",
                                enabled: true
                           }
                       ]
                   }
               ]
            }
        ]
    });
}
```

## <a name="localizing-the-json-blob"></a><span data-ttu-id="aa046-245">JSON BLOB のローカライズ</span><span class="sxs-lookup"><span data-stu-id="aa046-245">Localizing the JSON blob</span></span>

<span data-ttu-id="aa046-246">渡される JSON BLOB `requestCreateControls` は、カスタム コア タブのマニフェスト マークアップがローカライズされるのと同じ方法でローカライズされません ( [これは、「マニフェストからのコントロールのローカライズ](../develop/localization.md#control-localization-from-the-manifest)」で説明しています)。</span><span class="sxs-lookup"><span data-stu-id="aa046-246">The JSON blob that is passed to `requestCreateControls` is not localized the same way that the manifest markup for custom core tabs is localized (which is described at [Control localization from the manifest](../develop/localization.md#control-localization-from-the-manifest)).</span></span> <span data-ttu-id="aa046-247">代わりに、ロケールごとに異なる JSON BLOB を使用して、実行時にローカリゼーションを実行する必要があります。</span><span class="sxs-lookup"><span data-stu-id="aa046-247">Instead, the localization must occur at runtime using distinct JSON blobs for each locale.</span></span> <span data-ttu-id="aa046-248">`switch` [Office.context.displayLanguage](/javascript/api/office/office.context#displayLanguage)プロパティをテストするステートメントを使用することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="aa046-248">We suggest that you use a `switch` statement that tests the [Office.context.displayLanguage](/javascript/api/office/office.context#displayLanguage) property.</span></span> <span data-ttu-id="aa046-249">例を次に示します。</span><span class="sxs-lookup"><span data-stu-id="aa046-249">The following is an example:</span></span>

```javascript
function GetContextualTabsJsonSupportedLocale () {
    var displayLanguage = Office.context.displayLanguage;

        switch (displayLanguage) {
            case 'en-US':
                return `{
                    "actions": [
                        // actions omitted
                     ],
                    "tabs": [
                        {
                          "id": "CtxTab1",
                          "label": "Contoso Data",
                          "groups": [
                              // groups omitted
                          ]
                        }
                    ]
                }`;

            case 'fr-FR':
                return `{
                    "actions": [
                        // actions omitted 
                    ],
                    "tabs": [
                        {
                          "id": "CtxTab1",
                          "label": "Contoso Données",
                          "groups": [
                              // groups omitted
                          ]
                       }
                    ]
               }`;

            // Other cases omitted
       }
}
```

<span data-ttu-id="aa046-250">次に、次の例のように、コードは関数を呼び出して、 に渡されるローカライズされた BLOB `requestCreateControls` を取得します。</span><span class="sxs-lookup"><span data-stu-id="aa046-250">Then your code calls the function to get the localized blob that is passed to `requestCreateControls`, as in the following example:</span></span>

```javascript
var contextualTabJSON = GetContextualTabsJsonSupportedLocale();
```

## <a name="best-practices-for-custom-contextual-tabs"></a><span data-ttu-id="aa046-251">カスタム コンテキスト タブのベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="aa046-251">Best practices for custom contextual tabs</span></span>

### <a name="implement-an-alternate-ui-experience-when-custom-contextual-tabs-are-not-supported"></a><span data-ttu-id="aa046-252">カスタム コンテキスト タブがサポートされていない場合に、代替 UI エクスペリエンスを実装する</span><span class="sxs-lookup"><span data-stu-id="aa046-252">Implement an alternate UI experience when custom contextual tabs are not supported</span></span>

<span data-ttu-id="aa046-253">プラットフォーム、Office アプリケーション、およびOffice ビルドの一部の組み合わせでは、 がサポートしていません `requestCreateControls` 。</span><span class="sxs-lookup"><span data-stu-id="aa046-253">Some combinations of platform, Office application, and Office build don't support `requestCreateControls`.</span></span> <span data-ttu-id="aa046-254">アドインは、これらの組み合わせのいずれかでアドインを実行しているユーザーに代替エクスペリエンスを提供するように設計する必要があります。</span><span class="sxs-lookup"><span data-stu-id="aa046-254">Your add-in should be designed to provide an alternate experience to users who are running the add-in on one of those combinations.</span></span> <span data-ttu-id="aa046-255">次のセクションでは、フォールバック エクスペリエンスを提供する 2 つの方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="aa046-255">The following sections describe two ways of providing a fallback experience.</span></span>

#### <a name="use-noncontextual-tabs-or-controls"></a><span data-ttu-id="aa046-256">非コンテキスト タブまたはコントロールを使用する</span><span class="sxs-lookup"><span data-stu-id="aa046-256">Use noncontextual tabs or controls</span></span>

<span data-ttu-id="aa046-257">カスタム コンテキスト タブをサポートしないアプリケーションまたはプラットフォームでアドインが実行されている場合に、カスタム コンテキスト タブを実装するアドインでフォールバック エクスペリエンスを作成するように設計されたマニフェスト要素 [OverriddenByRibbonApi](../reference/manifest/overriddenbyribbonapi.md)があります。</span><span class="sxs-lookup"><span data-stu-id="aa046-257">There is a manifest element, [OverriddenByRibbonApi](../reference/manifest/overriddenbyribbonapi.md), that is designed to create a fallback experience in an add-in that implements custom contextual tabs when the add-in is running on an application or platform that doesn't support custom contextual tabs.</span></span> 

<span data-ttu-id="aa046-258">この要素を使用する最も簡単な方法は、アドインのカスタム コンテキスト タブのリボンのカスタマイズを複製する 1 つ以上のカスタム コア タブ ( *つまり、非コンテキスト* カスタム タブ) をマニフェストで定義することです。</span><span class="sxs-lookup"><span data-stu-id="aa046-258">The simplest strategy for using this element is that you define in the manifest one or more custom core tabs (that is, *noncontextual* custom tabs) that duplicate the ribbon customizations of the custom contextual tabs in your add-in.</span></span> <span data-ttu-id="aa046-259">しかし、 `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` [あなたは CustomTab](../reference/manifest/customtab.md)の最初の子要素として追加します。</span><span class="sxs-lookup"><span data-stu-id="aa046-259">But you add `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` as the first child element of the [CustomTab](../reference/manifest/customtab.md).</span></span> <span data-ttu-id="aa046-260">その結果、次のようになります。</span><span class="sxs-lookup"><span data-stu-id="aa046-260">The effect of doing so is the following:</span></span>

- <span data-ttu-id="aa046-261">カスタム コンテキスト タブをサポートするアプリケーションとプラットフォームでアドインを実行する場合、カスタム コア タブはリボンに表示されません。</span><span class="sxs-lookup"><span data-stu-id="aa046-261">If the add-in runs on an application and platform that support custom contextual tabs, then the custom core tab won't appear on the ribbon.</span></span> <span data-ttu-id="aa046-262">代わりに、アドインがメソッドを呼び出したときにカスタム コンテキスト タブが作成されます `requestCreateControls` 。</span><span class="sxs-lookup"><span data-stu-id="aa046-262">Instead, the custom contextual tab will be created when the add-in calls the `requestCreateControls` method.</span></span>
- <span data-ttu-id="aa046-263">アドインが をサポートしていないアプリケーションまたはプラットフォームで実行されている場合 *、* カスタム `requestCreateControls` コア タブがリボンに表示されます。</span><span class="sxs-lookup"><span data-stu-id="aa046-263">If the add-in runs on an application or platform that *doesn't* support `requestCreateControls`, then the custom core tab does appear on the ribbon.</span></span>

<span data-ttu-id="aa046-264">この単純な戦略の例を次に示します。</span><span class="sxs-lookup"><span data-stu-id="aa046-264">The following is an example of this simple strategy.</span></span>

```xml
<OfficeApp ...>
  ...
  <VersionOverrides ...>
    ...
    <Hosts>
      <Host ...>
        ...
        <DesktopFormFactor>
          <ExtensionPoint ...>
            <CustomTab ...>
              <OverriddenByRibbonApi>true</OverriddenByRibbonApi>
              ...
              <Group ...>
                ...
                <Control ... id="MyButton">
                  ...
                  <Action ...>
...
</OfficeApp>
```

<span data-ttu-id="aa046-265">この単純な戦略では、カスタム コンテキスト タブを子グループとコントロールと共に反映するカスタム コア タブを使用しますが、より複雑な戦略を使用できます。</span><span class="sxs-lookup"><span data-stu-id="aa046-265">This simple strategy uses a custom core tab that mirrors a custom contextual tab with it's child groups and controls, but you can use a more complex strategy.</span></span> <span data-ttu-id="aa046-266">`<OverriddenByRibbonApi>`要素は[、(](../reference/manifest/group.md)最初の) 子要素として 、グループ要素および[コントロール](../reference/manifest/control.md)要素 ([ボタンの種類](../reference/manifest/control.md#button-control)と[メニューの種類](../reference/manifest/control.md#menu-dropdown-button-controls)の両方) およびメニュー要素に追加することもできます `<Item>` 。</span><span class="sxs-lookup"><span data-stu-id="aa046-266">The `<OverriddenByRibbonApi>` element can also be added as (the first) child element to the [Group](../reference/manifest/group.md) and [Control](../reference/manifest/control.md) elements (both [button type](../reference/manifest/control.md#button-control) and [menu type](../reference/manifest/control.md#menu-dropdown-button-controls)), and menu `<Item>` elements.</span></span> <span data-ttu-id="aa046-267">この事実により、コンテキスト タブに表示されるグループやコントロールを、さまざまなカスタム コア タブのグループ、ボタン、メニューに分散させることができます。</span><span class="sxs-lookup"><span data-stu-id="aa046-267">This fact enables you to distribute the groups and controls that would otherwise appear on the contextual tab among various groups, buttons, and menus in various custom core tabs.</span></span> <span data-ttu-id="aa046-268">次に例を示します。</span><span class="sxs-lookup"><span data-stu-id="aa046-268">The following is an example.</span></span> <span data-ttu-id="aa046-269">カスタム コンテキスト タブがサポートされていない場合にのみ、カスタム コア タブに "MyButton" が表示されることに注意してください。</span><span class="sxs-lookup"><span data-stu-id="aa046-269">Note that "MyButton" will appear on the custom core tab only when custom contextual tabs are not supported.</span></span> <span data-ttu-id="aa046-270">ただし、カスタム コンテキスト タブがサポートされているかどうかに関係なく、親グループとカスタム コア タブが表示されます。</span><span class="sxs-lookup"><span data-stu-id="aa046-270">But the parent group and custom core tab will appear regardless of whether custom contextual tabs are supported.</span></span>

```xml
<OfficeApp ...>
  ...
  <VersionOverrides ...>
    ...
    <Hosts>
      <Host ...>
        ...
        <DesktopFormFactor>
          <ExtensionPoint ...>
            <CustomTab ...>              
              ...
              <Group ...>
                ...
                <Control ... id="MyButton">
                  <OverriddenByRibbonApi>true</OverriddenByRibbonApi>
                  ...
                  <Action ...>
...
</OfficeApp>
```

<span data-ttu-id="aa046-271">その他の例については、「 [オーバーライドされた ByRibbonApi](../reference/manifest/overriddenbyribbonapi.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="aa046-271">For more examples, see [OverriddenByRibbonApi](../reference/manifest/overriddenbyribbonapi.md).</span></span>

<span data-ttu-id="aa046-272">親タブ、グループ、またはメニューに マークが付いている場合、 `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` そのタブは表示されず、カスタム コンテキスト タブがサポートされていない場合は、子マークアップはすべて無視されます。</span><span class="sxs-lookup"><span data-stu-id="aa046-272">When a parent tab, group, or menu is marked with `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>`, then it isn't visible, and all of it's child markup is ignored, when custom contextual tabs aren't supported.</span></span> <span data-ttu-id="aa046-273">したがって、これらの子要素のいずれかが要素を持 `<OverriddenByRibbonApi>` っているかどうか、またはその値は関係ありません。</span><span class="sxs-lookup"><span data-stu-id="aa046-273">So, it doesn't matter if any of those child elements have the `<OverriddenByRibbonApi>` element or what its value is.</span></span> <span data-ttu-id="aa046-274">このことは、メニュー項目、コントロール、またはグループがすべてのコンテキストで表示される必要がある場合、そのメニュー項目、コントロール、またはグループをでマークしないだけでなく `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` *、その親メニュー、グループ、およびタブもこのようにマークしてはならない* ということです。</span><span class="sxs-lookup"><span data-stu-id="aa046-274">The implication of this is that if a menu item, control, or group must be visible in all contexts, then not only should it not be marked with `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>`, but *its ancestor menu, group, and tab must also not be marked this way*.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="aa046-275">タブ、グループ、または *メニューのすべての子* 要素を にマークを付けないでください `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` 。</span><span class="sxs-lookup"><span data-stu-id="aa046-275">Don't mark *all* of the child elements of a tab, group, or menu with `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>`.</span></span> <span data-ttu-id="aa046-276">前の段落で指定した理由で親要素にマークが付いている場合、これは無意味です `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` 。</span><span class="sxs-lookup"><span data-stu-id="aa046-276">This is pointless if the parent element is marked with `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` for reasons given in the preceding paragraph.</span></span> <span data-ttu-id="aa046-277">さらに、親の を除外する `<OverriddenByRibbonApi>` (または に設定 `false` する) 場合、カスタム コンテキスト タブがサポートされているかどうかに関係なく親が表示されますが、サポートされている場合は空になります。</span><span class="sxs-lookup"><span data-stu-id="aa046-277">Moreover, if you leave out the `<OverriddenByRibbonApi>` on the parent (or set it to `false`), then the parent will appear regardless of whether custom contextual tabs are supported, but it will be empty when they are supported.</span></span> <span data-ttu-id="aa046-278">したがって、カスタム コンテキスト タブがサポートされているときにすべての子要素が表示されない場合は、親と親のみを `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` でマークします。</span><span class="sxs-lookup"><span data-stu-id="aa046-278">So, if all the child elements shouldn't appear when custom contextual tabs are supported, mark the parent, and only the parent, with `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>`.</span></span>

#### <a name="use-apis-that-show-or-hide-a-task-pane-in-specified-contexts"></a><span data-ttu-id="aa046-279">指定したコンテキストで作業ウィンドウの表示と非表示を切り替える API を使用する</span><span class="sxs-lookup"><span data-stu-id="aa046-279">Use APIs that show or hide a task pane in specified contexts</span></span>

<span data-ttu-id="aa046-280">アドインの代わりに `<OverriddenByRibbonApi>` 、カスタム コンテキスト タブのコントロールの機能を複製する UI コントロールを含む作業ウィンドウを定義できます。次に[、Office.addin.showAsTaskpane](/javascript/api/office/office.addin?view=common-js&preserve-view=true#showAsTaskpane__)メソッドと[Office.addin.hide](/javascript/api/office/office.addin?view=common-js&preserve-view=true#hide__)メソッドを使用して、サポートされている場合にコンテキスト タブが表示された場合、およびコンテキスト タブが表示された場合にのみ、作業ウィンドウを表示します。</span><span class="sxs-lookup"><span data-stu-id="aa046-280">As an alternative to `<OverriddenByRibbonApi>`, your add-in can define a task pane with UI controls that duplicate the functionality of the controls on a custom contextual tab. Then use the [Office.addin.showAsTaskpane](/javascript/api/office/office.addin?view=common-js&preserve-view=true#showAsTaskpane__) and [Office.addin.hide](/javascript/api/office/office.addin?view=common-js&preserve-view=true#hide__) methods to show the task pane when, and only when, the contextual tab would have been shown if it was supported.</span></span> <span data-ttu-id="aa046-281">これらの方法の詳細については、「 [Office アドインの作業ウィンドウの表示と非表示を切り替える](../develop/show-hide-add-in.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="aa046-281">For details on how to use these methods, see [Show or hide the task pane of your Office Add-in](../develop/show-hide-add-in.md).</span></span>

### <a name="handle-the-hostrestartneeded-error"></a><span data-ttu-id="aa046-282">ホスト再起動が必要なエラーを処理します。</span><span class="sxs-lookup"><span data-stu-id="aa046-282">Handle the HostRestartNeeded error</span></span>

<span data-ttu-id="aa046-283">一部のシナリオでは、Office はリボンを更新できず、エラーを返します。</span><span class="sxs-lookup"><span data-stu-id="aa046-283">In some scenarios, Office is unable to update the ribbon and will return an error.</span></span> <span data-ttu-id="aa046-284">たとえば、アドインがアップグレードされ、アップグレードされたアドインに異なるカスタム アドイン コマンドのセットがある場合は、Office アプリケーションを閉じてから、もう一度開く必要があります。</span><span class="sxs-lookup"><span data-stu-id="aa046-284">For example, if the add-in is upgraded and the upgraded add-in has a different set of custom add-in commands, then the Office application must be closed and reopened.</span></span> <span data-ttu-id="aa046-285">それまでの間、`requestUpdate` メソッドは `HostRestartNeeded` エラーを返します。</span><span class="sxs-lookup"><span data-stu-id="aa046-285">Until it is, the `requestUpdate` method will return the error `HostRestartNeeded`.</span></span> <span data-ttu-id="aa046-286">コードでこのエラーを処理する必要があります。</span><span class="sxs-lookup"><span data-stu-id="aa046-286">Your code should handle this error.</span></span> <span data-ttu-id="aa046-287">以下は、その方法の例です。</span><span class="sxs-lookup"><span data-stu-id="aa046-287">The following is an example of how.</span></span> <span data-ttu-id="aa046-288">この場合、`reportError` メソッドがユーザーにエラーを表示します。</span><span class="sxs-lookup"><span data-stu-id="aa046-288">In this case, the `reportError` method displays the error to the user.</span></span>

```javascript
function showDataTab() {
    try {
        Office.ribbon.requestUpdate({
            tabs: [
                {
                    id: "CtxTab1",
                    visible: true
                }
            ]});
    }
    catch(error) {
        if (error.code == "HostRestartNeeded"){
            reportError("Contoso Awesome Add-in has been upgraded. Please save your work, then close and reopen the Office application.");
        }
    }
}
```
