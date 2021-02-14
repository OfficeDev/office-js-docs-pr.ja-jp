---
title: カスタム コンテキスト タブをアドインOffice作成する
description: カスタム コンテキスト タブをアドインに追加するOffice説明します。
ms.date: 01/29/2021
localization_priority: Normal
ms.openlocfilehash: 0badd779f22edc9b4659908409764bea1cde44f5
ms.sourcegitcommit: ccc0a86d099ab4f5ef3d482e4ae447c3f9b818a3
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/14/2021
ms.locfileid: "50237722"
---
# <a name="create-custom-contextual-tabs-in-office-add-ins-preview"></a><span data-ttu-id="76f2b-103">Office アドインでカスタム コンテキスト タブを作成する (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="76f2b-103">Create custom contextual tabs in Office Add-ins (preview)</span></span>

<span data-ttu-id="76f2b-104">操作依存タブは、指定したイベントがドキュメントで発生した場合にタブ行に表示される Office リボンの非表示のタブ コントロールOfficeします。</span><span class="sxs-lookup"><span data-stu-id="76f2b-104">A contextual tab is a hidden tab control in the Office ribbon that is displayed in the tab row when a specified event occurs in the Office document.</span></span> <span data-ttu-id="76f2b-105">たとえば、テーブルが **選択されている** ときに Excel リボンに表示される [テーブルのデザイン] タブです。</span><span class="sxs-lookup"><span data-stu-id="76f2b-105">For example, the **Table Design** tab that appears on the Excel ribbon when a table is selected.</span></span> <span data-ttu-id="76f2b-106">可視性を変更するイベント ハンドラーを作成することで、Office アドインにカスタム コンテキスト タブを含め、いつ表示または非表示にするか指定できます。</span><span class="sxs-lookup"><span data-stu-id="76f2b-106">You can include custom contextual tabs in your Office Add-in and specify when they are visible or hidden, by creating event handlers that change the visibility.</span></span> <span data-ttu-id="76f2b-107">(ただし、カスタム コンテキスト タブはフォーカスの変更には応答しない)。</span><span class="sxs-lookup"><span data-stu-id="76f2b-107">(However, custom contextual tabs do not respond to focus changes.)</span></span>

> [!NOTE]
> <span data-ttu-id="76f2b-108">この記事は、以下のドキュメントについて既に理解していることを前提としています。</span><span class="sxs-lookup"><span data-stu-id="76f2b-108">This article assumes that you are familiar with the following documentation.</span></span> <span data-ttu-id="76f2b-109">最近、アドイン コマンド (カスタム メニュー項目とリボン ボタン) を使用してない場合は、ドキュメントをご確認ください。</span><span class="sxs-lookup"><span data-stu-id="76f2b-109">Please review it if you haven't worked with Add-in Commands (custom menu items and ribbon buttons) recently.</span></span>
>
> - [<span data-ttu-id="76f2b-110">アドイン コマンドの基本概念</span><span class="sxs-lookup"><span data-stu-id="76f2b-110">Basic concepts for Add-in Commands</span></span>](add-in-commands.md)

> [!IMPORTANT]
> <span data-ttu-id="76f2b-111">カスタム コンテキスト タブはプレビュー中です。</span><span class="sxs-lookup"><span data-stu-id="76f2b-111">Custom contextual tabs are in preview.</span></span> <span data-ttu-id="76f2b-112">開発環境またはテスト環境で実験してください。ただし、実稼働アドインには追加しません。</span><span class="sxs-lookup"><span data-stu-id="76f2b-112">Please experiment with them in a development or testing environment but don't add them to a production add-in.</span></span>
>
> <span data-ttu-id="76f2b-113">カスタム コンテキスト タブは現在 Excel でのみサポートされ、次のプラットフォームとビルドでのみサポートされています。</span><span class="sxs-lookup"><span data-stu-id="76f2b-113">Custom contextual tabs are currently only supported on Excel and only on these platforms and builds:</span></span>
>
> - <span data-ttu-id="76f2b-114">Excel on Windows (永続的なライセンスではなく、Microsoft 365 のみ): バージョン 2011 (ビルド 13426.20274)。</span><span class="sxs-lookup"><span data-stu-id="76f2b-114">Excel on Windows (Microsoft 365 only, not perpetual license): Version 2011 (Build 13426.20274).</span></span> <span data-ttu-id="76f2b-115">Microsoft 365 サブスクリプションは、以前は 「月次チャネル (対象指定)」または「Insider Slow」と呼ばばされている現在のチャネル [(プレビュー)](https://insider.office.com/join/windows) に登録する必要がある場合があります。</span><span class="sxs-lookup"><span data-stu-id="76f2b-115">Your Microsoft 365 subscription may need to be on the [Current Channel (Preview)](https://insider.office.com/join/windows) formerly called "Monthly Channel (Targeted)" or "Insider Slow".</span></span>

> [!NOTE]
> <span data-ttu-id="76f2b-116">カスタム コンテキスト タブは、次の要件セットをサポートするプラットフォームでのみ動作します。</span><span class="sxs-lookup"><span data-stu-id="76f2b-116">Custom contextual tabs work only on platforms that support the following requirement sets.</span></span> <span data-ttu-id="76f2b-117">要件セットとそれらを使用する方法の詳細については、「アプリケーションと API の要件Office指定する」 [を参照してください](../develop/specify-office-hosts-and-api-requirements.md)。</span><span class="sxs-lookup"><span data-stu-id="76f2b-117">For more about requirement sets and how to work with them, see [Specify Office applications and API requirements](../develop/specify-office-hosts-and-api-requirements.md).</span></span>
>
> - [<span data-ttu-id="76f2b-118">SharedRuntime 1.1</span><span class="sxs-lookup"><span data-stu-id="76f2b-118">SharedRuntime 1.1</span></span>](../reference/requirement-sets/shared-runtime-requirement-sets.md)

## <a name="behavior-of-custom-contextual-tabs"></a><span data-ttu-id="76f2b-119">カスタム コンテキスト タブの動作</span><span class="sxs-lookup"><span data-stu-id="76f2b-119">Behavior of custom contextual tabs</span></span>

<span data-ttu-id="76f2b-120">カスタム コンテキスト タブのユーザー エクスペリエンスは、組み込みのコンテキスト タブのパターンOfficeに従います。</span><span class="sxs-lookup"><span data-stu-id="76f2b-120">The user experience for custom contextual tabs follows the pattern of built-in Office contextual tabs.</span></span> <span data-ttu-id="76f2b-121">配置カスタム コンテキスト タブの基本的な原則を次に示します。</span><span class="sxs-lookup"><span data-stu-id="76f2b-121">The following are the basic principles for the placement custom contextual tabs:</span></span>

- <span data-ttu-id="76f2b-122">カスタム コンテキスト タブが表示されている場合は、リボンの右側に表示されます。</span><span class="sxs-lookup"><span data-stu-id="76f2b-122">When a custom contextual tab is visible, it appears on the right end of the ribbon.</span></span>
- <span data-ttu-id="76f2b-123">1 つ以上の組み込みのコンテキスト タブと、アドインの 1 つ以上のカスタム コンテキスト タブが同時に表示される場合、カスタム コンテキスト タブは常にすべての組み込みコンテキスト タブの右側に表示されます。</span><span class="sxs-lookup"><span data-stu-id="76f2b-123">If one or more built-in contextual tabs and one or more custom contextual tabs from add-ins are visible at the same time, the custom contextual tabs are always to the right of all of the built-in contextual tabs.</span></span>
- <span data-ttu-id="76f2b-124">アドインに複数のコンテキスト タブがある場合に、複数のコンテキストが表示されている場合は、アドインで定義されている順序で表示されます。</span><span class="sxs-lookup"><span data-stu-id="76f2b-124">If your add-in has more than one contextual tab and there are contexts in which more than one is visible, they appear in the order in which they are defined in your add-in.</span></span> <span data-ttu-id="76f2b-125">(方向は Office 言語と同じ方向です。つまり、左から右の言語では左から右に、右から左の言語では右から左です)。定義 [方法の詳細については、「](#define-the-groups-and-controls-that-appear-on-the-tab) タブに表示されるグループとコントロールの定義」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="76f2b-125">(The direction is the same direction as the Office language; that is, is left-to-right in left-to-right languages, but right-to-left in right-to-left languages.) See [Define the groups and controls that appear on the tab](#define-the-groups-and-controls-that-appear-on-the-tab) for details about how you define them.</span></span>
- <span data-ttu-id="76f2b-126">特定のコンテキストで表示されるコンテキスト タブが複数のアドインにある場合は、アドインが起動された順序で表示されます。</span><span class="sxs-lookup"><span data-stu-id="76f2b-126">If more than one add-in has a contextual tab that is visible in a specific context, then they appear in the order in which the add-ins were launched.</span></span>
- <span data-ttu-id="76f2b-127">カスタム *コンテキスト* タブは、カスタム コア タブとは異なり、アプリケーションのリボンに完全Office追加されません。</span><span class="sxs-lookup"><span data-stu-id="76f2b-127">Custom *contextual* tabs, unlike custom core tabs, are not added permanently to the Office application's ribbon.</span></span> <span data-ttu-id="76f2b-128">アドインが実行されているOfficeドキュメントにのみ存在します。</span><span class="sxs-lookup"><span data-stu-id="76f2b-128">They are present only in Office documents on which your add-in is running.</span></span>

## <a name="major-steps-for-including-a-contextual-tab-in-an-add-in"></a><span data-ttu-id="76f2b-129">アドインにコンテキスト タブを含む主な手順</span><span class="sxs-lookup"><span data-stu-id="76f2b-129">Major steps for including a contextual tab in an add-in</span></span>

<span data-ttu-id="76f2b-130">アドインにカスタム コンテキスト タブを含む主な手順を次に示します。</span><span class="sxs-lookup"><span data-stu-id="76f2b-130">The following are the major steps for including a custom contextual tab in an add-in:</span></span>

1. <span data-ttu-id="76f2b-131">共有ランタイムを使用するアドインを構成します。</span><span class="sxs-lookup"><span data-stu-id="76f2b-131">Configure the add-in to use a shared runtime.</span></span>
1. <span data-ttu-id="76f2b-132">タブと、タブに表示されるグループとコントロールを定義します。</span><span class="sxs-lookup"><span data-stu-id="76f2b-132">Define the tab and the groups and controls that appear on it.</span></span>
1. <span data-ttu-id="76f2b-133">コンテキスト タブをユーザー設定にOffice。</span><span class="sxs-lookup"><span data-stu-id="76f2b-133">Register the contextual tab with Office.</span></span>
1. <span data-ttu-id="76f2b-134">タブが表示される状況を指定します。</span><span class="sxs-lookup"><span data-stu-id="76f2b-134">Specify the circumstances when the tab will be visible.</span></span>

## <a name="configure-the-add-in-to-use-a-shared-runtime"></a><span data-ttu-id="76f2b-135">共有ランタイムを使用するアドインを構成する</span><span class="sxs-lookup"><span data-stu-id="76f2b-135">Configure the add-in to use a shared runtime</span></span>

<span data-ttu-id="76f2b-136">カスタム コンテキスト タブを追加するには、アドインで共有ランタイムを使用する必要があります。</span><span class="sxs-lookup"><span data-stu-id="76f2b-136">Adding custom contextual tabs requires your add-in to use the shared runtime.</span></span> <span data-ttu-id="76f2b-137">詳細については、「共有ランタイム [を使用するアドインを構成する」を参照してください](../develop/configure-your-add-in-to-use-a-shared-runtime.md)。</span><span class="sxs-lookup"><span data-stu-id="76f2b-137">For more information, see [Configure an add-in to use a shared runtime](../develop/configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

## <a name="define-the-groups-and-controls-that-appear-on-the-tab"></a><span data-ttu-id="76f2b-138">タブに表示されるグループとコントロールを定義する</span><span class="sxs-lookup"><span data-stu-id="76f2b-138">Define the groups and controls that appear on the tab</span></span>

<span data-ttu-id="76f2b-139">マニフェスト内の XML で定義されたカスタム コア タブとは異なり、カスタム コンテキスト タブは実行時に JSON BLOB を使用して定義されます。</span><span class="sxs-lookup"><span data-stu-id="76f2b-139">Unlike custom core tabs, which are defined with XML in the manifest, custom contextual tabs are defined at runtime with a JSON blob.</span></span> <span data-ttu-id="76f2b-140">コードは BLOB を JavaScript オブジェクトに解析し、そのオブジェクトを [Office.ribbon.requestCreateControls メソッドに渡](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls-tabDefinition-) します。</span><span class="sxs-lookup"><span data-stu-id="76f2b-140">Your code parses the blob into a JavaScript object, and then passes the object to the [Office.ribbon.requestCreateControls](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls-tabDefinition-) method.</span></span> <span data-ttu-id="76f2b-141">カスタム コンテキスト タブは、アドインが現在実行されているドキュメントにのみ表示されます。</span><span class="sxs-lookup"><span data-stu-id="76f2b-141">Custom contextual tabs are only present in documents on which your add-in is currently running.</span></span> <span data-ttu-id="76f2b-142">これは、アドインのインストール時に Office アプリケーション リボンに追加され、別のドキュメントが開かれたときに存在し続けるカスタム コア タブとは異なります。</span><span class="sxs-lookup"><span data-stu-id="76f2b-142">This is different from custom core tabs which are added to the Office application ribbon when the add-in is installed and remain present when another document is opened.</span></span> <span data-ttu-id="76f2b-143">また、 `requestCreateControls` このメソッドはアドインのセッションで 1 回だけ実行できます。</span><span class="sxs-lookup"><span data-stu-id="76f2b-143">Also, the `requestCreateControls` method can be run only once in a session of your add-in.</span></span> <span data-ttu-id="76f2b-144">再度呼び出された場合は、エラーがスローされます。</span><span class="sxs-lookup"><span data-stu-id="76f2b-144">If it is called again, an error is thrown.</span></span>

> [!NOTE]
> <span data-ttu-id="76f2b-145">JSON BLOB のプロパティとサブプロパティ (およびキー名) の構造は [、CustomTab](../reference/manifest/customtab.md) 要素とそのマニフェスト XML 内の子孫要素の構造と大まかに平行です。</span><span class="sxs-lookup"><span data-stu-id="76f2b-145">The structure of the JSON blob's properties and subproperties (and the key names) is roughly parallel to the structure of the [CustomTab](../reference/manifest/customtab.md) element and its descendant elements in the manifest XML.</span></span>

<span data-ttu-id="76f2b-146">コンテキスト タブ JSON BLOB のステップ バイ ステップで例を作成します。</span><span class="sxs-lookup"><span data-stu-id="76f2b-146">We'll construct an example of a contextual tabs JSON blob step-by-step.</span></span> <span data-ttu-id="76f2b-147">(コンテキスト タブ JSON の完全なスキーマは、dynamic-ribbon.schema.js[ です](https://developer.microsoft.com/json-schemas/office-js/dynamic-ribbon.schema.json)。</span><span class="sxs-lookup"><span data-stu-id="76f2b-147">(The full schema for the contextual tab JSON is at [dynamic-ribbon.schema.json](https://developer.microsoft.com/json-schemas/office-js/dynamic-ribbon.schema.json).</span></span> <span data-ttu-id="76f2b-148">このリンクは、コンテキスト タブのプレビュー期間中に機能しない可能性があります。</span><span class="sxs-lookup"><span data-stu-id="76f2b-148">This link may not be working in the preview period for contextual tabs.</span></span> <span data-ttu-id="76f2b-149">リンクが機能しない場合は、下書きページでスキーマの最新の下書 [きdynamic-ribbon.schema.jsを見つける必要があります](https://github.com/OfficeDev/testing-assets/tree/master/jsonschema/dynamic-ribbon/1.0/dynamic-ribbon.schema.json)。コードで作業している場合Visual Studioこのファイルを使用して、JSON IntelliSenseを取得し、検証できます。</span><span class="sxs-lookup"><span data-stu-id="76f2b-149">If the link is not working, you can find the latest draft of the schema at [draft dynamic-ribbon.schema.json](https://github.com/OfficeDev/testing-assets/tree/master/jsonschema/dynamic-ribbon/1.0/dynamic-ribbon.schema.json).) If you are working in Visual Studio Code, you can use this file to get IntelliSense and to validate your JSON.</span></span> <span data-ttu-id="76f2b-150">詳細については、「コード - JSON スキーマと [設定を使用Visual Studio JSON](https://code.visualstudio.com/docs/languages/json#_json-schemas-and-settings)の編集」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="76f2b-150">For more information, see [Editing JSON with Visual Studio Code - JSON schemas and settings](https://code.visualstudio.com/docs/languages/json#_json-schemas-and-settings).</span></span>


1. <span data-ttu-id="76f2b-151">まず、次の 2 つの配列プロパティを持つ JSON 文字列を作成 `actions` します `tabs` 。</span><span class="sxs-lookup"><span data-stu-id="76f2b-151">Begin by creating a JSON string with two array properties named `actions` and `tabs`.</span></span> <span data-ttu-id="76f2b-152">配列 `actions` は、操作別タブのコントロールで実行できるすべての関数の仕様です。配列 `tabs` は、最大 *20* までの 1 つ以上のコンテキスト タブを定義します。</span><span class="sxs-lookup"><span data-stu-id="76f2b-152">The `actions` array is a specification of all the functions that can be executed by controls on the contextual tab. The `tabs` array defines one or more contextual tabs, *up to a maximum of 20*.</span></span>

    ```json
    '{
      "actions": [

      ],
      "tabs": [

      ]
    }'
    ```

1. <span data-ttu-id="76f2b-153">この操作別タブの単純な例にはボタンが 1 つしか含めなく、したがってアクションは 1 つのみです。</span><span class="sxs-lookup"><span data-stu-id="76f2b-153">This simple example of a contextual tab will have only a single button and, thus, only a single action.</span></span> <span data-ttu-id="76f2b-154">以下を配列の唯一のメンバーとして追加 `actions` します。</span><span class="sxs-lookup"><span data-stu-id="76f2b-154">Add the following as the only member of the `actions` array.</span></span> <span data-ttu-id="76f2b-155">このマークアップについては、次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="76f2b-155">About this markup, note:</span></span>

    - <span data-ttu-id="76f2b-156">プロパティ `id` `type` とプロパティは必須です。</span><span class="sxs-lookup"><span data-stu-id="76f2b-156">The `id` and `type` properties are mandatory.</span></span>
    - <span data-ttu-id="76f2b-157">値には `type` 、"ExecuteFunction" または "ShowTaskpane" を指定できます。</span><span class="sxs-lookup"><span data-stu-id="76f2b-157">The value of `type` can be either "ExecuteFunction" or "ShowTaskpane".</span></span>
    - <span data-ttu-id="76f2b-158">プロパティ `functionName` は、値が次の場合にのみ使用 `type` されます `ExecuteFunction` 。</span><span class="sxs-lookup"><span data-stu-id="76f2b-158">The `functionName` property is only used when the value of `type` is `ExecuteFunction`.</span></span> <span data-ttu-id="76f2b-159">FunctionFile で定義されている関数の名前です。</span><span class="sxs-lookup"><span data-stu-id="76f2b-159">It is the name of a function defined in the FunctionFile.</span></span> <span data-ttu-id="76f2b-160">FunctionFile の詳細については、「アドイン コマンドの基本 [概念」を参照してください](add-in-commands.md)。</span><span class="sxs-lookup"><span data-stu-id="76f2b-160">For more information about the FunctionFile, see [Basic concepts for Add-in Commands](add-in-commands.md).</span></span>
    - <span data-ttu-id="76f2b-161">後の手順では、このアクションをコンテキスト タブのボタンにマップします。</span><span class="sxs-lookup"><span data-stu-id="76f2b-161">In a later step, you will map this action to a button on the contextual tab.</span></span>

    ```json
    {
      "id": "executeWriteData",
      "type": "ExecuteFunction",
      "functionName": "writeData"
    }
   ```

1. <span data-ttu-id="76f2b-162">以下を配列の唯一のメンバーとして追加 `tabs` します。</span><span class="sxs-lookup"><span data-stu-id="76f2b-162">Add the following as the only member of the `tabs` array.</span></span> <span data-ttu-id="76f2b-163">このマークアップについては、次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="76f2b-163">About this markup, note:</span></span>

    - <span data-ttu-id="76f2b-164">`id` プロパティは必須です。</span><span class="sxs-lookup"><span data-stu-id="76f2b-164">The `id` property is required.</span></span> <span data-ttu-id="76f2b-165">アドイン内のすべてのコンテキスト タブの中で一意である簡潔でわかりやすい ID を使用します。</span><span class="sxs-lookup"><span data-stu-id="76f2b-165">Use a brief, descriptive ID that is unique among all contextual tabs in your add-in.</span></span>
    - <span data-ttu-id="76f2b-166">`label` プロパティは必須です。</span><span class="sxs-lookup"><span data-stu-id="76f2b-166">The `label` property is required.</span></span> <span data-ttu-id="76f2b-167">コンテキスト タブのラベルとして使用すると、ユーザーに分け親しまれる文字列になります。</span><span class="sxs-lookup"><span data-stu-id="76f2b-167">It is a user-friendly string to serve as the label of the contextual tab.</span></span>
    - <span data-ttu-id="76f2b-168">`groups` プロパティは必須です。</span><span class="sxs-lookup"><span data-stu-id="76f2b-168">The `groups` property is required.</span></span> <span data-ttu-id="76f2b-169">タブに表示されるコントロールのグループを定義します。少なくとも 1 つのメンバーと *20 以下である必要があります*。</span><span class="sxs-lookup"><span data-stu-id="76f2b-169">It defines the groups of controls that will appear on the tab. It must have at least one member *and no more than 20*.</span></span> <span data-ttu-id="76f2b-170">(カスタム コンテキスト タブに設定できるコントロールの数にも制限があります。また、持っているグループの数も制限されます。</span><span class="sxs-lookup"><span data-stu-id="76f2b-170">(There are also limits on the number of controls that you can have on a custom contextual tab and that will also constrain how many groups that you have.</span></span> <span data-ttu-id="76f2b-171">詳細については、次の手順を参照してください)。</span><span class="sxs-lookup"><span data-stu-id="76f2b-171">See the next step for more information.)</span></span>

    > [!NOTE]
    > <span data-ttu-id="76f2b-172">タブ オブジェクトには、アドインの起動直後にタブを表示するかどうかを指定するオプションのプロパティ `visible` を指定することもできます。</span><span class="sxs-lookup"><span data-stu-id="76f2b-172">The tab object can also have an optional `visible` property that specifies whether the tab is visible immediately when the add-in starts up.</span></span> <span data-ttu-id="76f2b-173">コンテキスト タブは通常、ユーザー イベントが表示をトリガーするまで非表示になります (ユーザーがドキュメント内の何らかの種類のエンティティを選択した場合など)、プロパティは既定で存在しない場合に設定されます `visible` `false` 。</span><span class="sxs-lookup"><span data-stu-id="76f2b-173">Since contextual tabs are normally hidden until a user event triggers their visibility (such as the user selecting an entity of some type in the document), the `visible` property defaults to `false` when not present.</span></span> <span data-ttu-id="76f2b-174">後のセクションでは、イベントに応答してプロパティを `true` 設定する方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="76f2b-174">In a later section, we show how to set the property to `true` in response to an event.</span></span>

    ```json
    {
      "id": "CtxTab1",
      "label": "Contoso Data",
      "groups": [

      ]
    }
    ```

1. <span data-ttu-id="76f2b-175">単純な例では、コンテキスト タブには 1 つのグループのみがあります。</span><span class="sxs-lookup"><span data-stu-id="76f2b-175">In the simple ongoing example, the contextual tab has only a single group.</span></span> <span data-ttu-id="76f2b-176">以下を配列の唯一のメンバーとして追加 `groups` します。</span><span class="sxs-lookup"><span data-stu-id="76f2b-176">Add the following as the only member of the `groups` array.</span></span> <span data-ttu-id="76f2b-177">このマークアップについては、次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="76f2b-177">About this markup, note:</span></span>

    - <span data-ttu-id="76f2b-178">すべてのプロパティが必要です。</span><span class="sxs-lookup"><span data-stu-id="76f2b-178">All the properties are required.</span></span>
    - <span data-ttu-id="76f2b-179">この `id` プロパティは、タブ内のすべてのグループ間で一意である必要があります。簡潔でわかりやすい ID を使用します。</span><span class="sxs-lookup"><span data-stu-id="76f2b-179">The `id` property must be unique among all the groups in the tab. Use a brief, descriptive ID.</span></span>
    - <span data-ttu-id="76f2b-180">グループ `label` のラベルとして使用する、ユーザー に分かしい文字列です。</span><span class="sxs-lookup"><span data-stu-id="76f2b-180">The `label` is a user-friendly string to serve as the label of the group.</span></span>
    - <span data-ttu-id="76f2b-181">プロパティの値は、リボンのサイズとアプリケーション ウィンドウのサイズに応じてリボンに表示されるアイコンを指定するOffice `icon` 配列です。</span><span class="sxs-lookup"><span data-stu-id="76f2b-181">The `icon` property's value is an array of objects that specify the icons that the group will have on the ribbon depending on the size of the ribbon and the Office application window.</span></span>
    - <span data-ttu-id="76f2b-182">プロパティ `controls` の値は、グループ内のボタンとメニューを指定するオブジェクトの配列です。</span><span class="sxs-lookup"><span data-stu-id="76f2b-182">The `controls` property's value is an array of objects that specify the buttons and menus in the group.</span></span> <span data-ttu-id="76f2b-183">少なくとも 1 つが必要です。</span><span class="sxs-lookup"><span data-stu-id="76f2b-183">There must be at least one.</span></span>

    > [!IMPORTANT]
    > <span data-ttu-id="76f2b-184">*タブ全体のコントロールの総数は 20 以下です。*</span><span class="sxs-lookup"><span data-stu-id="76f2b-184">*The total number of controls on the whole tab can be no more than 20.*</span></span> <span data-ttu-id="76f2b-185">たとえば、それぞれ 6 つのコントロールを持つ 3 つのグループと 2 つのコントロールを持つ 4 つ目のグループを持つ場合がありますが、4 つのグループにそれぞれ 6 つのコントロールを持つすることはできません。</span><span class="sxs-lookup"><span data-stu-id="76f2b-185">For example, you could have 3 groups with 6 controls each, and a fourth group with 2 controls, but you cannot have 4 groups with 6 controls each.</span></span>  

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

1. <span data-ttu-id="76f2b-186">すべてのグループには、32x32 px と 80x80 px の 2 つ以上のサイズのアイコンが必要です。</span><span class="sxs-lookup"><span data-stu-id="76f2b-186">Every group must have an icon of at least two sizes, 32x32 px and 80x80 px.</span></span> <span data-ttu-id="76f2b-187">必要に応じて、16x16 px、20x20 px、24x24 px、40x40 px、48x48 px、64x64 px のアイコンを設定できます。</span><span class="sxs-lookup"><span data-stu-id="76f2b-187">Optionally, you can also have icons of sizes 16x16 px, 20x20 px, 24x24 px, 40x40 px, 48x48 px, and 64x64 px.</span></span> <span data-ttu-id="76f2b-188">Office、リボンとアプリケーション ウィンドウのサイズに基づいて使用するアイコンOffice決定します。</span><span class="sxs-lookup"><span data-stu-id="76f2b-188">Office decides which icon to use based on the size of the ribbon and Office application window.</span></span> <span data-ttu-id="76f2b-189">アイコン配列に次のオブジェクトを追加します。</span><span class="sxs-lookup"><span data-stu-id="76f2b-189">Add the following objects to the icon array.</span></span> <span data-ttu-id="76f2b-190">(ウィンドウとリボンのサイズが、グループ上のコントロールの少なくとも1 つが表示されるのに十分な大きさの場合、グループ アイコンは表示されません。</span><span class="sxs-lookup"><span data-stu-id="76f2b-190">(If the window and ribbon sizes are large enough for at least one of the *controls* on the group to appear, then no group icon at all appears.</span></span> <span data-ttu-id="76f2b-191">たとえば、Word **ウィンドウを縮小** して展開する場合は、Word リボンの [スタイル] グループを参照してください)。このマークアップについては、次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="76f2b-191">For an example, watch the **Styles** group on the Word ribbon as you shrink and expand the Word window.) About this markup, note:</span></span>

    - <span data-ttu-id="76f2b-192">両方のプロパティが必要です。</span><span class="sxs-lookup"><span data-stu-id="76f2b-192">Both the properties are required.</span></span>
    - <span data-ttu-id="76f2b-193">プロパティ `size` の単位はピクセルです。</span><span class="sxs-lookup"><span data-stu-id="76f2b-193">The `size` property unit of measure is pixels.</span></span> <span data-ttu-id="76f2b-194">アイコンは常に正方形なので、数値は高さと幅の両方です。</span><span class="sxs-lookup"><span data-stu-id="76f2b-194">Icons are always square, so the number is both the height and the width.</span></span>
    - <span data-ttu-id="76f2b-195">この `sourceLocation` プロパティは、アイコンの完全な URL を指定します。</span><span class="sxs-lookup"><span data-stu-id="76f2b-195">The `sourceLocation` property specifies the full URL to the icon.</span></span>

    > [!IMPORTANT]
    > <span data-ttu-id="76f2b-196">開発から実稼働に移行する場合 (ドメインを localhost から contoso.com に変更する場合など) アドインのマニフェストの URL を通常は変更する必要があるのと同様に、コンテキスト タブ JSON の URL も変更する必要があります。</span><span class="sxs-lookup"><span data-stu-id="76f2b-196">Just as you typically must change the URLs in the add-in's manifest when you move from development to production (such as changing the domain from localhost to contoso.com), you must also change the URLs in your contextual tabs JSON.</span></span>

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

1. <span data-ttu-id="76f2b-197">この単純な例では、グループにボタンが 1 つしか表示されます。</span><span class="sxs-lookup"><span data-stu-id="76f2b-197">In our simple ongoing example, the group has only a single button.</span></span> <span data-ttu-id="76f2b-198">次のオブジェクトを配列の唯一のメンバーとして追加 `controls` します。</span><span class="sxs-lookup"><span data-stu-id="76f2b-198">Add the following object as the only member of the `controls` array.</span></span> <span data-ttu-id="76f2b-199">このマークアップについては、次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="76f2b-199">About this markup, note:</span></span>

    - <span data-ttu-id="76f2b-200">ただし、すべてのプロパティ `enabled` は必須です。</span><span class="sxs-lookup"><span data-stu-id="76f2b-200">All the properties, except `enabled`, are required.</span></span>
    - <span data-ttu-id="76f2b-201">`type` コントロールの種類を指定します。</span><span class="sxs-lookup"><span data-stu-id="76f2b-201">`type` specifies the type of control.</span></span> <span data-ttu-id="76f2b-202">値には、"Button"、"Menu"、または "MobileButton" を指定できます。</span><span class="sxs-lookup"><span data-stu-id="76f2b-202">The values can be "Button", "Menu", or "MobileButton".</span></span>
    - <span data-ttu-id="76f2b-203">`id` 最大 125 文字まで入力できます。</span><span class="sxs-lookup"><span data-stu-id="76f2b-203">`id` can be up to 125 characters.</span></span> 
    - <span data-ttu-id="76f2b-204">`actionId` は、配列で定義されているアクションの ID である必要 `actions` があります。</span><span class="sxs-lookup"><span data-stu-id="76f2b-204">`actionId` must be the ID of an action defined in the `actions` array.</span></span> <span data-ttu-id="76f2b-205">(このセクションの手順 1 を参照してください)。</span><span class="sxs-lookup"><span data-stu-id="76f2b-205">(See step 1 of this section.)</span></span>
    - <span data-ttu-id="76f2b-206">`label` は、ボタンのラベルとして使用する、ユーザー に使い分け可能な文字列です。</span><span class="sxs-lookup"><span data-stu-id="76f2b-206">`label` is a user-friendly string to serve as the label of the button.</span></span>
    - <span data-ttu-id="76f2b-207">`superTip` は、豊富な形式のツール ヒントを表します。</span><span class="sxs-lookup"><span data-stu-id="76f2b-207">`superTip` represents a rich form of tool tip.</span></span> <span data-ttu-id="76f2b-208">プロパティと `title` プロパティ `description` の両方が必要です。</span><span class="sxs-lookup"><span data-stu-id="76f2b-208">Both the `title` and `description` properties are required.</span></span>
    - <span data-ttu-id="76f2b-209">`icon` ボタンのアイコンを指定します。</span><span class="sxs-lookup"><span data-stu-id="76f2b-209">`icon` specifies the icons for the button.</span></span> <span data-ttu-id="76f2b-210">グループ アイコンに関する前の注釈もここに適用されます。</span><span class="sxs-lookup"><span data-stu-id="76f2b-210">The previous remarks about the group icon apply here too.</span></span>
    - <span data-ttu-id="76f2b-211">`enabled` (オプション) コンテキスト タブが表示される際にボタンを有効にするかどうかを指定します。</span><span class="sxs-lookup"><span data-stu-id="76f2b-211">`enabled` (optional) specifies whether the button is enabled when the contextual tab appears starts up.</span></span> <span data-ttu-id="76f2b-212">存在しない場合の既定値は次の値です `true` 。</span><span class="sxs-lookup"><span data-stu-id="76f2b-212">The default if not present is `true`.</span></span> 

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
 
<span data-ttu-id="76f2b-213">JSON BLOB の完全な例を次に示します。</span><span class="sxs-lookup"><span data-stu-id="76f2b-213">The following is the complete example of the JSON blob:</span></span>

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

## <a name="register-the-contextual-tab-with-office-with-requestcreatecontrols"></a><span data-ttu-id="76f2b-214">requestCreateControls を使用してOfficeタブを登録する</span><span class="sxs-lookup"><span data-stu-id="76f2b-214">Register the contextual tab with Office with requestCreateControls</span></span>

<span data-ttu-id="76f2b-215">コンテキスト タブは [、Office.ribbon.requestCreateControls](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls_tabDefinition_) メソッドをOfficeして、コンテキスト タブに登録されます。</span><span class="sxs-lookup"><span data-stu-id="76f2b-215">The contextual tab is registered with Office by calling the [Office.ribbon.requestCreateControls](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls_tabDefinition_) method.</span></span> <span data-ttu-id="76f2b-216">これは通常、メソッドに割り当てられている関数またはメソッドで `Office.initialize` 行 `Office.onReady` われます。</span><span class="sxs-lookup"><span data-stu-id="76f2b-216">This is typically done in either the function that is assigned to `Office.initialize` or with the `Office.onReady` method.</span></span> <span data-ttu-id="76f2b-217">これらのメソッドとアドインの初期化の詳細については、「アドインの初期化Office [参照してください](../develop/initialize-add-in.md)。</span><span class="sxs-lookup"><span data-stu-id="76f2b-217">For more about these methods and initializing the add-in, see [Initialize your Office Add-in](../develop/initialize-add-in.md).</span></span> <span data-ttu-id="76f2b-218">ただし、初期化後はメソッドを呼び出す必要があります。</span><span class="sxs-lookup"><span data-stu-id="76f2b-218">You can, however, call the method anytime after initialization.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="76f2b-219">この `requestCreateControls` メソッドは、アドインの特定のセッションで 1 回だけ呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="76f2b-219">The `requestCreateControls` method can be called only once in a given session of an add-in.</span></span> <span data-ttu-id="76f2b-220">再度呼び出された場合は、エラーがスローされます。</span><span class="sxs-lookup"><span data-stu-id="76f2b-220">An error is thrown if it is called again.</span></span>

<span data-ttu-id="76f2b-221">次に例を示します。</span><span class="sxs-lookup"><span data-stu-id="76f2b-221">The following is an example.</span></span> <span data-ttu-id="76f2b-222">JSON 文字列を JavaScript 関数に渡す前に、メソッドを使用して `JSON.parse` JavaScript オブジェクトに変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="76f2b-222">Note that the JSON string must be converted to a JavaScript object with the `JSON.parse` method before it can be passed to a JavaScript function.</span></span>

```javascript
Office.onReady(async () => {
    const contextualTabJSON = ` ... `; // Assign the JSON string such as the one at the end of the preceding section.
    const contextualTab = JSON.parse(contextualTabJSON);
    await Office.ribbon.requestCreateControls(contextualTab);
});
```

## <a name="specify-the-contexts-when-the-tab-will-be-visible-with-requestupdate"></a><span data-ttu-id="76f2b-223">requestUpdate でタブが表示されるコンテキストを指定する</span><span class="sxs-lookup"><span data-stu-id="76f2b-223">Specify the contexts when the tab will be visible with requestUpdate</span></span>

<span data-ttu-id="76f2b-224">通常、カスタム コンテキスト タブは、ユーザーが開始するイベントによってアドインのコンテキストが変更されると表示されます。</span><span class="sxs-lookup"><span data-stu-id="76f2b-224">Typically, a custom contextual tab should appear when a user-initiated event changes the add-in context.</span></span> <span data-ttu-id="76f2b-225">(Excel ブックの既定のワークシートにある) グラフがアクティブ化されている場合にのみ、タブが表示されるシナリオを考えます。</span><span class="sxs-lookup"><span data-stu-id="76f2b-225">Consider a scenario in which the tab should be visible when, and only when, a chart (on the default worksheet of an Excel workbook) is activated.</span></span>

<span data-ttu-id="76f2b-226">まず、ハンドラーを割り当てる必要があります。</span><span class="sxs-lookup"><span data-stu-id="76f2b-226">Begin by assigning handlers.</span></span> <span data-ttu-id="76f2b-227">これは通常、次の例のようにメソッドで行われます。この例では、ハンドラー (後の手順で作成) をワークシート内のすべてのグラフのイベントに割り当 `Office.onReady` `onActivated` `onDeactivated` てる必要があります。</span><span class="sxs-lookup"><span data-stu-id="76f2b-227">This is commonly done in the `Office.onReady` method as in the following example which assigns handlers (created in a later step) to the `onActivated` and `onDeactivated` events of all the charts in the worksheet.</span></span>

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

<span data-ttu-id="76f2b-228">次に、ハンドラーを定義します。</span><span class="sxs-lookup"><span data-stu-id="76f2b-228">Next, define the handlers.</span></span> <span data-ttu-id="76f2b-229">次に示すのは単純な例ですが、より堅牢なバージョンの関数については、この記事で後の `showDataTab` [「HostRestartNeeded](#handle-the-hostrestartneeded-error) エラーの処理」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="76f2b-229">The following is a simple example of a `showDataTab`, but see [Handling the HostRestartNeeded error](#handle-the-hostrestartneeded-error) later in this article for a more robust version of the function.</span></span> <span data-ttu-id="76f2b-230">このコードについては、以下の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="76f2b-230">About this code, note:</span></span>

- <span data-ttu-id="76f2b-231">Office では、リボンの状態を更新するタイミングが制御されます。</span><span class="sxs-lookup"><span data-stu-id="76f2b-231">Office controls when it updates the state of the ribbon.</span></span> <span data-ttu-id="76f2b-232">[Office.ribbon.requestUpdate](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestupdate-input-)メソッドは、更新要求をキューに入れられます。</span><span class="sxs-lookup"><span data-stu-id="76f2b-232">The  [Office.ribbon.requestUpdate](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestupdate-input-) method queues a request to update.</span></span> <span data-ttu-id="76f2b-233">このメソッドは、リボンが実際に更新されるのではなく、要求をキューに入れ次第、オブジェクト `Promise` を解決します。</span><span class="sxs-lookup"><span data-stu-id="76f2b-233">The method will resolve the `Promise` object as soon as it has queued the request, not when the ribbon actually updates.</span></span>
- <span data-ttu-id="76f2b-234">メソッドのパラメーターは `requestUpdate` [RibbonUpdaterData](/javascript/api/office/office.ribbonupdaterdata) オブジェクトで、(1) *JSON* で指定されているとおりに ID でタブを指定し、(2) タブの可視性を指定します。</span><span class="sxs-lookup"><span data-stu-id="76f2b-234">The parameter for the `requestUpdate` method is a [RibbonUpdaterData](/javascript/api/office/office.ribbonupdaterdata) object that (1) specifies the tab by its ID *exactly as specified in the JSON* and (2) specifies visibility of the tab.</span></span>
- <span data-ttu-id="76f2b-235">同じコンテキストで表示するカスタム コンテキスト タブが複数ある場合は、単純にタブ オブジェクトを配列に追加 `tabs` します。</span><span class="sxs-lookup"><span data-stu-id="76f2b-235">If you have more than one custom contextual tab that should be visible in the same context, you simply add additional tab objects to the `tabs` array.</span></span>

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

<span data-ttu-id="76f2b-236">タブを非表示にするハンドラーは、プロパティを設定し戻す以外は、ほぼ `visible` 同じです `false` 。</span><span class="sxs-lookup"><span data-stu-id="76f2b-236">The handler to hide the tab is nearly identical, except that it sets the `visible` property back to `false`.</span></span>

<span data-ttu-id="76f2b-237">またOffice JavaScript ライブラリには、オブジェクトの作成を容易にするためのインターフェイス (型) `RibbonUpdateData` がいくつか用意されています。</span><span class="sxs-lookup"><span data-stu-id="76f2b-237">The Office JavaScript library also provides several interfaces (types) to make it easier to construct the`RibbonUpdateData` object.</span></span> <span data-ttu-id="76f2b-238">TypeScript の `showDataTab` 関数を次に示します。この関数は、これらの型を利用します。</span><span class="sxs-lookup"><span data-stu-id="76f2b-238">The following is the `showDataTab` function in TypeScript and it makes use of these types.</span></span>

```typescript
const showDataTab = async () => {
    const myContextualTab: Office.Tab = {id: "CtxTab1", visible: true};
    const ribbonUpdater: Office.RibbonUpdaterData = { tabs: [ myContextualTab ]};
    await Office.ribbon.requestUpdate(ribbonUpdater);
}
```

### <a name="toggle-tab-visibility-and-the-enabled-status-of-a-button-at-the-same-time"></a><span data-ttu-id="76f2b-239">タブの表示とボタンの有効な状態を同時に切り替える</span><span class="sxs-lookup"><span data-stu-id="76f2b-239">Toggle tab visibility and the enabled status of a button at the same time</span></span>

<span data-ttu-id="76f2b-240">このメソッドは、カスタム コンテキスト タブまたはカスタム コア タブのカスタム ボタンの有効または無効の状態を切り替 `requestUpdate` える場合にも使用されます。詳細については、「アドイン コマンドを [有効または無効にする」を参照してください](disable-add-in-commands.md)。</span><span class="sxs-lookup"><span data-stu-id="76f2b-240">The `requestUpdate` method is also used to toggle the enabled or disabled status of a custom button on either a custom contextual tab or a custom core tab. For details about this, see [Enable and Disable Add-in Commands](disable-add-in-commands.md).</span></span> <span data-ttu-id="76f2b-241">タブの表示とボタンの有効な状態の両方を同時に変更するシナリオが考えられます。</span><span class="sxs-lookup"><span data-stu-id="76f2b-241">There may be scenarios in which you want to change both the visibility of a tab and the enabled status of a button at the same time.</span></span> <span data-ttu-id="76f2b-242">これは、1 回の呼び出しで行います `requestUpdate` 。</span><span class="sxs-lookup"><span data-stu-id="76f2b-242">You can do this with a single call of `requestUpdate`.</span></span> <span data-ttu-id="76f2b-243">次の例では、コンテキスト タブが表示されるのと同時に、コア タブのボタンが有効になります。</span><span class="sxs-lookup"><span data-stu-id="76f2b-243">The following is an example in which a button on a core tab is enabled at the same time as a contextual tab is made visible.</span></span>

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

<span data-ttu-id="76f2b-244">次の例では、有効になっているボタンは、表示されているのと同じコンテキスト タブにあります。</span><span class="sxs-lookup"><span data-stu-id="76f2b-244">In the following example, the button that is enabled is on the very same contextual tab that is being made visible.</span></span>

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

## <a name="localizing-the-json-blob"></a><span data-ttu-id="76f2b-245">JSON BLOB のローカライズ</span><span class="sxs-lookup"><span data-stu-id="76f2b-245">Localizing the JSON blob</span></span>

<span data-ttu-id="76f2b-246">渡される JSON BLOB は、カスタム コア タブのマニフェスト マークアップがローカライズされるのと同じ方法ではローカライズされません (マニフェストからのコントロールのローカライズで `requestCreateControls` [説明します](../develop/localization.md#control-localization-from-the-manifest))。</span><span class="sxs-lookup"><span data-stu-id="76f2b-246">The JSON blob that is passed to `requestCreateControls` is not localized the same way that the manifest markup for custom core tabs is localized (which is described at [Control localization from the manifest](../develop/localization.md#control-localization-from-the-manifest)).</span></span> <span data-ttu-id="76f2b-247">代わりに、ローカライズは、ロケールごとに異なる JSON BLOB を使用して実行時に行う必要があります。</span><span class="sxs-lookup"><span data-stu-id="76f2b-247">Instead, the localization must occur at runtime using distinct JSON blobs for each locale.</span></span> <span data-ttu-id="76f2b-248">`switch` [Office.context.displayLanguage](/javascript/api/office/office.context#displayLanguage)プロパティをテストするステートメントを使用してください。</span><span class="sxs-lookup"><span data-stu-id="76f2b-248">We suggest that you use a `switch` statement that tests the [Office.context.displayLanguage](/javascript/api/office/office.context#displayLanguage) property.</span></span> <span data-ttu-id="76f2b-249">例を次に示します。</span><span class="sxs-lookup"><span data-stu-id="76f2b-249">The following is an example:</span></span>

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

<span data-ttu-id="76f2b-250">次に、次の例のように、コードで関数を呼び出して、渡されるローカライズされた BLOB `requestCreateControls` を取得します。</span><span class="sxs-lookup"><span data-stu-id="76f2b-250">Then your code calls the function to get the localized blob that is passed to `requestCreateControls`, as in the following example:</span></span>

```javascript
var contextualTabJSON = GetContextualTabsJsonSupportedLocale();
```

## <a name="best-practices-for-custom-contextual-tabs"></a><span data-ttu-id="76f2b-251">カスタム コンテキスト タブのベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="76f2b-251">Best practices for custom contextual tabs</span></span>

### <a name="implement-an-alternate-ui-experience-when-custom-contextual-tabs-are-not-supported"></a><span data-ttu-id="76f2b-252">カスタム コンテキスト タブがサポートされていない場合に代替 UI エクスペリエンスを実装する</span><span class="sxs-lookup"><span data-stu-id="76f2b-252">Implement an alternate UI experience when custom contextual tabs are not supported</span></span>

<span data-ttu-id="76f2b-253">プラットフォーム、アプリケーション、Office、およびOfficeの組み合わせはサポートされていません `requestCreateControls` 。</span><span class="sxs-lookup"><span data-stu-id="76f2b-253">Some combinations of platform, Office application, and Office build don't support `requestCreateControls`.</span></span> <span data-ttu-id="76f2b-254">アドインは、これらの組み合わせの 1 つでアドインを実行しているユーザーに代替エクスペリエンスを提供するように設計する必要があります。</span><span class="sxs-lookup"><span data-stu-id="76f2b-254">Your add-in should be designed to provide an alternate experience to users who are running the add-in on one of those combinations.</span></span> <span data-ttu-id="76f2b-255">次のセクションでは、フォールバック エクスペリエンスを提供する 2 つの方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="76f2b-255">The following sections describe two ways of providing a fallback experience.</span></span>

#### <a name="use-noncontextual-tabs-or-controls"></a><span data-ttu-id="76f2b-256">コンテキスト以外のタブまたはコントロールを使用する</span><span class="sxs-lookup"><span data-stu-id="76f2b-256">Use noncontextual tabs or controls</span></span>

<span data-ttu-id="76f2b-257">マニフェスト要素 [OverriddenByRibbonApi](../reference/manifest/overriddenbyribbonapi.md)は、カスタム コンテキスト タブをサポートしないアプリケーションまたはプラットフォームでアドインが実行されている場合に、カスタム コンテキスト タブを実装するアドインでフォールバック エクスペリエンスを作成するように設計されています。</span><span class="sxs-lookup"><span data-stu-id="76f2b-257">There is a manifest element, [OverriddenByRibbonApi](../reference/manifest/overriddenbyribbonapi.md), that is designed to create a fallback experience in an add-in that implements custom contextual tabs when the add-in is running on an application or platform that doesn't support custom contextual tabs.</span></span> 

<span data-ttu-id="76f2b-258">この要素を使用する最も簡単な方法は、アドインのカスタム コンテキスト タブのリボンカスタマイズを複製する 1 つ以上のカスタム コア タブ (つまり、非コンテキスト カスタム タブ) をマニフェストで定義する方法です。</span><span class="sxs-lookup"><span data-stu-id="76f2b-258">The simplest strategy for using this element is that you define in the manifest one or more custom core tabs (that is, *noncontextual* custom tabs) that duplicate the ribbon customizations of the custom contextual tabs in your add-in.</span></span> <span data-ttu-id="76f2b-259">ただし `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` [、CustomTab](../reference/manifest/customtab.md)の最初の子要素として追加します。</span><span class="sxs-lookup"><span data-stu-id="76f2b-259">But you add `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` as the first child element of the [CustomTab](../reference/manifest/customtab.md).</span></span> <span data-ttu-id="76f2b-260">その効果は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="76f2b-260">The effect of doing so is the following:</span></span>

- <span data-ttu-id="76f2b-261">カスタム コンテキスト タブをサポートするアプリケーションとプラットフォームでアドインを実行する場合、カスタムコア タブはリボンに表示されません。</span><span class="sxs-lookup"><span data-stu-id="76f2b-261">If the add-in runs on an application and platform that support custom contextual tabs, then the custom core tab won't appear on the ribbon.</span></span> <span data-ttu-id="76f2b-262">代わりに、アドインがメソッドを呼び出す際にカスタム コンテキスト タブが作成 `requestCreateControls` されます。</span><span class="sxs-lookup"><span data-stu-id="76f2b-262">Instead, the custom contextual tab will be created when the add-in calls the `requestCreateControls` method.</span></span>
- <span data-ttu-id="76f2b-263">アドインがサポートしていないアプリケーションまたはプラットフォームで実行される場合、カスタム コア `requestCreateControls` タブがリボンに表示されます。</span><span class="sxs-lookup"><span data-stu-id="76f2b-263">If the add-in runs on an application or platform that *doesn't* support `requestCreateControls`, then the custom core tab does appear on the ribbon.</span></span>

<span data-ttu-id="76f2b-264">この簡単な戦略の例を次に示します。</span><span class="sxs-lookup"><span data-stu-id="76f2b-264">The following is an example of this simple strategy.</span></span>

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

<span data-ttu-id="76f2b-265">この簡単な戦略では、カスタム コンテキスト タブと子グループとコントロールをミラー化するカスタム コア タブを使用しますが、より複雑な戦略を使用できます。</span><span class="sxs-lookup"><span data-stu-id="76f2b-265">This simple strategy uses a custom core tab that mirrors a custom contextual tab with it's child groups and controls, but you can use a more complex strategy.</span></span> <span data-ttu-id="76f2b-266">要素は、Group 要素と Control 要素 (ボタンの種類とメニューの種類の両方) とメニュー要素に (最初の) 子要素として `<OverriddenByRibbonApi>` 追加[](../reference/manifest/control.md#button-control)[](../reference/manifest/group.md)[](../reference/manifest/control.md)[](../reference/manifest/control.md#menu-dropdown-button-controls) `<Item>` することもできます。</span><span class="sxs-lookup"><span data-stu-id="76f2b-266">The `<OverriddenByRibbonApi>` element can also be added as (the first) child element to the [Group](../reference/manifest/group.md) and [Control](../reference/manifest/control.md) elements (both [button type](../reference/manifest/control.md#button-control) and [menu type](../reference/manifest/control.md#menu-dropdown-button-controls)), and menu `<Item>` elements.</span></span> <span data-ttu-id="76f2b-267">この事実により、コンテキスト タブに表示されるグループとコントロールを、さまざまなカスタム コア タブのさまざまなグループ、ボタン、メニューに分散できます。</span><span class="sxs-lookup"><span data-stu-id="76f2b-267">This fact enables you to distribute the groups and controls that would otherwise appear on the contextual tab among various groups, buttons, and menus in various custom core tabs.</span></span> <span data-ttu-id="76f2b-268">次に例を示します。</span><span class="sxs-lookup"><span data-stu-id="76f2b-268">The following is an example.</span></span> <span data-ttu-id="76f2b-269">"MyButton" は、カスタム コンテキスト タブがサポートされていない場合にのみ、カスタム コア タブに表示されます。</span><span class="sxs-lookup"><span data-stu-id="76f2b-269">Note that "MyButton" will appear on the custom core tab only when custom contextual tabs are not supported.</span></span> <span data-ttu-id="76f2b-270">ただし、カスタム コンテキスト タブがサポートされるかどうかに関係なく、親グループとカスタムコア タブが表示されます。</span><span class="sxs-lookup"><span data-stu-id="76f2b-270">But the parent group and custom core tab will appear regardless of whether custom contextual tabs are supported.</span></span>

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

<span data-ttu-id="76f2b-271">その他の例については [、「OverriddenByRibbonApi」を参照してください](../reference/manifest/overriddenbyribbonapi.md)。</span><span class="sxs-lookup"><span data-stu-id="76f2b-271">For more examples, see [OverriddenByRibbonApi](../reference/manifest/overriddenbyribbonapi.md).</span></span>

<span data-ttu-id="76f2b-272">親タブ、グループ、またはメニューにマークが付いている場合、そのタブは表示されません。カスタム コンテキスト タブがサポートされていない場合、そのすべての子マークアップは無視されます `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` 。</span><span class="sxs-lookup"><span data-stu-id="76f2b-272">When a parent tab, group, or menu is marked with `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>`, then it isn't visible, and all of it's child markup is ignored, when custom contextual tabs aren't supported.</span></span> <span data-ttu-id="76f2b-273">そのため、これらの子要素の中に要素がある場合や、その値 `<OverriddenByRibbonApi>` が何かは関係ありません。</span><span class="sxs-lookup"><span data-stu-id="76f2b-273">So, it doesn't matter if any of those child elements have the `<OverriddenByRibbonApi>` element or what its value is.</span></span> <span data-ttu-id="76f2b-274">これは、メニュー項目、コントロール、またはグループをすべてのコンテキストで表示する必要がある場合、メニュー項目、コントロール、またはグループがマークされていないだけでなく、その先祖のメニュー、グループ、およびタブもこの方法でマークされなければならないという意味です `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` 。 </span><span class="sxs-lookup"><span data-stu-id="76f2b-274">The implication of this is that if a menu item, control, or group must be visible in all contexts, then not only should it not be marked with `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>`, but *its ancestor menu, group, and tab must also not be marked this way*.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="76f2b-275">タブ、グループ *、または* メニューのすべての子要素にマークを付けはしない `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` 。</span><span class="sxs-lookup"><span data-stu-id="76f2b-275">Don't mark *all* of the child elements of a tab, group, or menu with `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>`.</span></span> <span data-ttu-id="76f2b-276">前の段落で説明した理由で親要素にマークが付いている場合、 `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` これは無意味です。</span><span class="sxs-lookup"><span data-stu-id="76f2b-276">This is pointless if the parent element is marked with `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` for reasons given in the preceding paragraph.</span></span> <span data-ttu-id="76f2b-277">さらに、親のタブを指定しない (または親に設定した) 場合は、カスタム コンテキスト タブがサポートされているかどうかに関係なく親が表示されますが、サポートされている場合は空になります。 `<OverriddenByRibbonApi>` `false`</span><span class="sxs-lookup"><span data-stu-id="76f2b-277">Moreover, if you leave out the `<OverriddenByRibbonApi>` on the parent (or set it to `false`), then the parent will appear regardless of whether custom contextual tabs are supported, but it will be empty when they are supported.</span></span> <span data-ttu-id="76f2b-278">したがって、カスタム コンテキスト タブがサポートされているときにすべての子要素が表示されない場合は、親と親のみをマークします `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` 。</span><span class="sxs-lookup"><span data-stu-id="76f2b-278">So, if all the child elements shouldn't appear when custom contextual tabs are supported, mark the parent, and only the parent, with `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>`.</span></span>

#### <a name="use-apis-that-show-or-hide-a-task-pane-in-specified-contexts"></a><span data-ttu-id="76f2b-279">指定したコンテキストで作業ウィンドウを表示または非表示にする API を使用する</span><span class="sxs-lookup"><span data-stu-id="76f2b-279">Use APIs that show or hide a task pane in specified contexts</span></span>

<span data-ttu-id="76f2b-280">代わりに、アドインは、カスタム コンテキスト タブのコントロールの機能を複製する UI コントロールを含む作業ウィンドウ `<OverriddenByRibbonApi>` を定義できます。 [次に、Office.addin.showAsTaskpane](/javascript/api/office/office.addin?view=common-js&preserve-view=true#showAsTaskpane__) メソッドと [Office.addin.hide](/javascript/api/office/office.addin?view=common-js&preserve-view=true#hide__) メソッドを使用して、操作別タブがサポートされている場合にのみ作業ウィンドウを表示します。</span><span class="sxs-lookup"><span data-stu-id="76f2b-280">As an alternative to `<OverriddenByRibbonApi>`, your add-in can define a task pane with UI controls that duplicate the functionality of the controls on a custom contextual tab. Then use the [Office.addin.showAsTaskpane](/javascript/api/office/office.addin?view=common-js&preserve-view=true#showAsTaskpane__) and [Office.addin.hide](/javascript/api/office/office.addin?view=common-js&preserve-view=true#hide__) methods to show the task pane when, and only when, the contextual tab would have been shown if it was supported.</span></span> <span data-ttu-id="76f2b-281">これらのメソッドの使い方の詳細については、アドインの作業ウィンドウを表示または非表示にするOffice [参照してください](../develop/show-hide-add-in.md)。</span><span class="sxs-lookup"><span data-stu-id="76f2b-281">For details on how to use these methods, see [Show or hide the task pane of your Office Add-in](../develop/show-hide-add-in.md).</span></span>

### <a name="handle-the-hostrestartneeded-error"></a><span data-ttu-id="76f2b-282">HostRestartNeeded エラーの処理</span><span class="sxs-lookup"><span data-stu-id="76f2b-282">Handle the HostRestartNeeded error</span></span>

<span data-ttu-id="76f2b-283">一部のシナリオでは、Office はリボンを更新できず、エラーを返します。</span><span class="sxs-lookup"><span data-stu-id="76f2b-283">In some scenarios, Office is unable to update the ribbon and will return an error.</span></span> <span data-ttu-id="76f2b-284">たとえば、アドインがアップグレードされ、アップグレードされたアドインに異なるカスタム アドイン コマンドのセットがある場合は、Office アプリケーションを閉じてから、もう一度開く必要があります。</span><span class="sxs-lookup"><span data-stu-id="76f2b-284">For example, if the add-in is upgraded and the upgraded add-in has a different set of custom add-in commands, then the Office application must be closed and reopened.</span></span> <span data-ttu-id="76f2b-285">それまでの間、`requestUpdate` メソッドは `HostRestartNeeded` エラーを返します。</span><span class="sxs-lookup"><span data-stu-id="76f2b-285">Until it is, the `requestUpdate` method will return the error `HostRestartNeeded`.</span></span> <span data-ttu-id="76f2b-286">コードでこのエラーを処理する必要があります。</span><span class="sxs-lookup"><span data-stu-id="76f2b-286">Your code should handle this error.</span></span> <span data-ttu-id="76f2b-287">その方法の例を次に示します。</span><span class="sxs-lookup"><span data-stu-id="76f2b-287">The following is an example of how.</span></span> <span data-ttu-id="76f2b-288">この場合、`reportError` メソッドがユーザーにエラーを表示します。</span><span class="sxs-lookup"><span data-stu-id="76f2b-288">In this case, the `reportError` method displays the error to the user.</span></span>

```javascript
function showDataTab() {
    try {
        await Office.ribbon.requestUpdate({
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
