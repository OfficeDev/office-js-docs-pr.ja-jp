---
title: カスタム コンテキスト タブをアドインOffice作成する
description: カスタム コンテキスト タブをアドインに追加するOffice説明します。
ms.date: 11/20/2020
localization_priority: Normal
ms.openlocfilehash: 3939e3338c734e1d6400dc261b59e35de63e5779
ms.sourcegitcommit: 545888b08f57bb1babb05ccfd83b2b3286bdad5c
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 01/08/2021
ms.locfileid: "49789136"
---
# <a name="create-custom-contextual-tabs-in-office-add-ins-preview"></a><span data-ttu-id="a3e82-103">アドインのカスタム コンテキスト タブOfficeする (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="a3e82-103">Create custom contextual tabs in Office Add-ins (preview)</span></span>

<span data-ttu-id="a3e82-104">操作依存タブは、指定したイベントがドキュメントで発生した場合にタブ行に表示される Office リボンの非表示のタブ コントロールOfficeします。</span><span class="sxs-lookup"><span data-stu-id="a3e82-104">A contextual tab is a hidden tab control in the Office ribbon that is displayed in the tab row when a specified event occurs in the Office document.</span></span> <span data-ttu-id="a3e82-105">たとえば、テーブルが **選択されている** ときに Excel リボンに表示される [テーブルのデザイン] タブです。</span><span class="sxs-lookup"><span data-stu-id="a3e82-105">For example, the **Table Design** tab that appears on the Excel ribbon when a table is selected.</span></span> <span data-ttu-id="a3e82-106">可視性を変更するイベント ハンドラーを作成することで、Office アドインにカスタム コンテキスト タブを含め、いつ表示または非表示にするか指定できます。</span><span class="sxs-lookup"><span data-stu-id="a3e82-106">You can include custom contextual tabs in your Office add-in and specify when they are visible or hidden, by creating event handlers that change the visibility.</span></span> <span data-ttu-id="a3e82-107">(ただし、カスタム コンテキスト タブはフォーカスの変更には応答しない)。</span><span class="sxs-lookup"><span data-stu-id="a3e82-107">(However, custom contextual tabs do not respond to focus changes.)</span></span>

> [!NOTE]
> <span data-ttu-id="a3e82-108">この記事は、以下のドキュメントについて既に理解していることを前提としています。</span><span class="sxs-lookup"><span data-stu-id="a3e82-108">This article assumes that you are familiar with the following documentation.</span></span> <span data-ttu-id="a3e82-109">最近、アドイン コマンド (カスタム メニュー項目とリボン ボタン) を使用してない場合は、ドキュメントをご確認ください。</span><span class="sxs-lookup"><span data-stu-id="a3e82-109">Please review it if you haven't worked with Add-in Commands (custom menu items and ribbon buttons) recently.</span></span>
>
> - [<span data-ttu-id="a3e82-110">アドイン コマンドの基本概念</span><span class="sxs-lookup"><span data-stu-id="a3e82-110">Basic concepts for Add-in Commands</span></span>](add-in-commands.md)

> [!IMPORTANT]
> <span data-ttu-id="a3e82-111">カスタム コンテキスト タブはプレビュー中です。</span><span class="sxs-lookup"><span data-stu-id="a3e82-111">Custom contextual tabs are in preview.</span></span> <span data-ttu-id="a3e82-112">開発環境またはテスト環境で実験してください。ただし、実稼働アドインには追加しません。</span><span class="sxs-lookup"><span data-stu-id="a3e82-112">Please experiment with them in a development or testing environment but don't add them to a production add-in.</span></span>
>
> <span data-ttu-id="a3e82-113">カスタム コンテキスト タブは現在 Excel でのみサポートされ、次のプラットフォームとビルドでのみサポートされています。</span><span class="sxs-lookup"><span data-stu-id="a3e82-113">Custom contextual tabs are currently only supported on Excel and only on these platforms and builds:</span></span>
>
> - <span data-ttu-id="a3e82-114">Excel on Windows (永続的なライセンスではなく、Microsoft 365 のみ): バージョン 2011 (ビルド 13426.20274)。</span><span class="sxs-lookup"><span data-stu-id="a3e82-114">Excel on Windows (Microsoft 365 only, not perpetual license): Version 2011 (Build 13426.20274).</span></span> <span data-ttu-id="a3e82-115">Microsoft 365 サブスクリプションは、以前は 「月次チャネル (対象指定)」または「Insider Slow」と呼ばばされている現在のチャネル [(プレビュー)](https://insider.office.com/join/windows) に登録する必要がある場合があります。</span><span class="sxs-lookup"><span data-stu-id="a3e82-115">Your Microsoft 365 subscription may need to be on the [Current Channel (Preview)](https://insider.office.com/join/windows) formerly called "Monthly Channel (Targeted)" or "Insider Slow".</span></span>

> [!NOTE]
> <span data-ttu-id="a3e82-116">カスタム コンテキスト タブは、次の要件セットをサポートするプラットフォームでのみ動作します。</span><span class="sxs-lookup"><span data-stu-id="a3e82-116">Custom contextual tabs work only on platforms that support the following requirement sets.</span></span> <span data-ttu-id="a3e82-117">要件セットとそれらを使用する方法の詳細については、「アプリケーションと API の要件Office指定する」 [を参照してください](../develop/specify-office-hosts-and-api-requirements.md)。</span><span class="sxs-lookup"><span data-stu-id="a3e82-117">For more about requirement sets and how to work with them, see [Specify Office applications and API requirements](../develop/specify-office-hosts-and-api-requirements.md).</span></span>
>
> - [<span data-ttu-id="a3e82-118">SharedRuntime 1.1</span><span class="sxs-lookup"><span data-stu-id="a3e82-118">SharedRuntime 1.1</span></span>](../reference/requirement-sets/shared-runtime-requirement-sets.md)

## <a name="behavior-of-custom-contextual-tabs"></a><span data-ttu-id="a3e82-119">カスタム コンテキスト タブの動作</span><span class="sxs-lookup"><span data-stu-id="a3e82-119">Behavior of custom contextual tabs</span></span>

<span data-ttu-id="a3e82-120">カスタム コンテキスト タブのユーザー エクスペリエンスは、組み込みのコンテキスト タブのパターンOfficeに従います。</span><span class="sxs-lookup"><span data-stu-id="a3e82-120">The user experience for custom contextual tabs follows the pattern of built-in Office contextual tabs.</span></span> <span data-ttu-id="a3e82-121">配置カスタム コンテキスト タブの基本的な原則を次に示します。</span><span class="sxs-lookup"><span data-stu-id="a3e82-121">The following are the basic principles for the placement custom contextual tabs:</span></span>

- <span data-ttu-id="a3e82-122">カスタム コンテキスト タブが表示されている場合は、リボンの右側に表示されます。</span><span class="sxs-lookup"><span data-stu-id="a3e82-122">When a custom contextual tab is visible, it appears on the right end of the ribbon.</span></span>
- <span data-ttu-id="a3e82-123">1 つ以上の組み込みのコンテキスト タブと、アドインの 1 つ以上のカスタム コンテキスト タブが同時に表示される場合、カスタム コンテキスト タブは常にすべての組み込みコンテキスト タブの右側に表示されます。</span><span class="sxs-lookup"><span data-stu-id="a3e82-123">If one or more built-in contextual tabs and one or more custom contextual tabs from add-ins are visible at the same time, the custom contextual tabs are always to the right of all of the built-in contextual tabs.</span></span>
- <span data-ttu-id="a3e82-124">アドインに複数のコンテキスト タブがある場合に、複数のコンテキストが表示されている場合は、アドインで定義されている順序で表示されます。</span><span class="sxs-lookup"><span data-stu-id="a3e82-124">If your add-in has more than one contextual tab and there are contexts in which more than one is visible, they appear in the order in which they are defined in your add-in.</span></span> <span data-ttu-id="a3e82-125">(方向は Office 言語と同じ方向です。つまり、左から右の言語では左から右に、右から左の言語では右から左です)。定義 [方法の詳細については、「](#define-the-groups-and-controls-that-appear-on-the-tab) タブに表示されるグループとコントロールの定義」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="a3e82-125">(The direction is the same direction as the Office language; that is, is left-to-right in left-to-right languages, but right-to-left in right-to-left languages.) See [Define the groups and controls that appear on the tab](#define-the-groups-and-controls-that-appear-on-the-tab) for details about how you define them.</span></span>
- <span data-ttu-id="a3e82-126">特定のコンテキストで表示されるコンテキスト タブが複数のアドインにある場合は、アドインが起動された順序で表示されます。</span><span class="sxs-lookup"><span data-stu-id="a3e82-126">If more than one add-in has a contextual tab that is visible in a specific context, then they appear in the order in which the add-ins were launched.</span></span>
- <span data-ttu-id="a3e82-127">カスタム *コンテキスト* タブは、カスタム コア タブとは異なり、アプリケーションのリボンに完全Office追加されません。</span><span class="sxs-lookup"><span data-stu-id="a3e82-127">Custom *contextual* tabs, unlike custom core tabs, are not added permanently to the Office application's ribbon.</span></span> <span data-ttu-id="a3e82-128">アドインが実行されているOfficeドキュメントにのみ存在します。</span><span class="sxs-lookup"><span data-stu-id="a3e82-128">They are present only in Office documents on which your add-in is running.</span></span>

## <a name="major-steps-for-including-a-contextual-tab-in-an-add-in"></a><span data-ttu-id="a3e82-129">アドインにコンテキスト タブを含む主な手順</span><span class="sxs-lookup"><span data-stu-id="a3e82-129">Major steps for including a contextual tab in an add-in</span></span>

<span data-ttu-id="a3e82-130">アドインにカスタム コンテキスト タブを含む主な手順を次に示します。</span><span class="sxs-lookup"><span data-stu-id="a3e82-130">The following are the major steps for including a custom contextual tab in an add-in:</span></span>

1. <span data-ttu-id="a3e82-131">共有ランタイムを使用するアドインを構成します。</span><span class="sxs-lookup"><span data-stu-id="a3e82-131">Configure the add-in to use a shared runtime.</span></span>
1. <span data-ttu-id="a3e82-132">タブと、タブに表示されるグループとコントロールを定義します。</span><span class="sxs-lookup"><span data-stu-id="a3e82-132">Define the tab and the groups and controls that appear on it.</span></span>
1. <span data-ttu-id="a3e82-133">コンテキスト タブをユーザー設定にOffice。</span><span class="sxs-lookup"><span data-stu-id="a3e82-133">Register the contextual tab with Office.</span></span>
1. <span data-ttu-id="a3e82-134">タブが表示される状況を指定します。</span><span class="sxs-lookup"><span data-stu-id="a3e82-134">Specify the circumstances when the tab will be visible.</span></span>

## <a name="configure-the-add-in-to-use-a-shared-runtime"></a><span data-ttu-id="a3e82-135">共有ランタイムを使用するアドインを構成する</span><span class="sxs-lookup"><span data-stu-id="a3e82-135">Configure the add-in to use a shared runtime</span></span>

<span data-ttu-id="a3e82-136">カスタム コンテキスト タブを追加するには、アドインで共有ランタイムを使用する必要があります。</span><span class="sxs-lookup"><span data-stu-id="a3e82-136">Adding custom contextual tabs requires your add-in to use the shared runtime.</span></span> <span data-ttu-id="a3e82-137">詳細については、「共有ランタイム [を使用するアドインを構成する」を参照してください](../develop/configure-your-add-in-to-use-a-shared-runtime.md)。</span><span class="sxs-lookup"><span data-stu-id="a3e82-137">For more information, see [Configure an add-in to use a shared runtime](../develop/configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

## <a name="define-the-groups-and-controls-that-appear-on-the-tab"></a><span data-ttu-id="a3e82-138">タブに表示されるグループとコントロールを定義する</span><span class="sxs-lookup"><span data-stu-id="a3e82-138">Define the groups and controls that appear on the tab</span></span>

<span data-ttu-id="a3e82-139">マニフェスト内の XML で定義されたカスタム コア タブとは異なり、カスタム コンテキスト タブは実行時に JSON BLOB を使用して定義されます。</span><span class="sxs-lookup"><span data-stu-id="a3e82-139">Unlike custom core tabs, which are defined with XML in the manifest, custom contextual tabs are defined at runtime with a JSON blob.</span></span> <span data-ttu-id="a3e82-140">コードは BLOB を JavaScript オブジェクトに解析し、そのオブジェクトを [Office.ribbon.requestCreateControls メソッドに渡](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls-tabDefinition-) します。</span><span class="sxs-lookup"><span data-stu-id="a3e82-140">Your code parses the blob into a JavaScript object, and then passes the object to the [Office.ribbon.requestCreateControls](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls-tabDefinition-) method.</span></span> <span data-ttu-id="a3e82-141">カスタム コンテキスト タブは、アドインが現在実行されているドキュメントにのみ表示されます。</span><span class="sxs-lookup"><span data-stu-id="a3e82-141">Custom contextual tabs are only present in documents on which your add-in is currently running.</span></span> <span data-ttu-id="a3e82-142">これは、アドインのインストール時に Office アプリケーション リボンに追加され、別のドキュメントが開かれたときに存在し続けるカスタム コア タブとは異なります。</span><span class="sxs-lookup"><span data-stu-id="a3e82-142">This is different from custom core tabs which are added to the Office application ribbon when the add-in is installed and remain present when another document is opened.</span></span> <span data-ttu-id="a3e82-143">また、 `requestCreateControls` このメソッドはアドインのセッションで 1 回だけ実行できます。</span><span class="sxs-lookup"><span data-stu-id="a3e82-143">Also, the `requestCreateControls` method can be run only once in a session of your add-in.</span></span> <span data-ttu-id="a3e82-144">再度呼び出された場合は、エラーがスローされます。</span><span class="sxs-lookup"><span data-stu-id="a3e82-144">If it is called again, an error is thrown.</span></span>

> [!NOTE]
> <span data-ttu-id="a3e82-145">JSON BLOB のプロパティとサブプロパティ (およびキー名) の構造は [、CustomTab](../reference/manifest/customtab.md) 要素とそのマニフェスト XML 内の子孫要素の構造と大まかに平行です。</span><span class="sxs-lookup"><span data-stu-id="a3e82-145">The structure of the JSON blob's properties and subproperties (and the key names) is roughly parallel to the structure of the [CustomTab](../reference/manifest/customtab.md) element and its descendant elements in the manifest XML.</span></span>

<span data-ttu-id="a3e82-146">コンテキスト タブ JSON BLOB のステップ バイ ステップで例を作成します。</span><span class="sxs-lookup"><span data-stu-id="a3e82-146">We'll construct an example of a contextual tabs JSON blob step-by-step.</span></span> <span data-ttu-id="a3e82-147">(コンテキスト タブ JSON の完全なスキーマは、dynamic-ribbon.schema.js[ です](https://developer.microsoft.com/json-schemas/office-js/dynamic-ribbon.schema.json)。</span><span class="sxs-lookup"><span data-stu-id="a3e82-147">(The full schema for the contextual tab JSON is at [dynamic-ribbon.schema.json](https://developer.microsoft.com/json-schemas/office-js/dynamic-ribbon.schema.json).</span></span> <span data-ttu-id="a3e82-148">このリンクは、コンテキスト タブのプレビュー期間の早い段階では機能しない可能性があります。</span><span class="sxs-lookup"><span data-stu-id="a3e82-148">This link may not be working in the early preview period for contextual tabs.</span></span> <span data-ttu-id="a3e82-149">リンクが機能しない場合は、下書きページでスキーマの最新の下書 [きdynamic-ribbon.schema.jsを見つける必要があります](https://github.com/OfficeDev/testing-assets/tree/master/jsonschema/dynamic-ribbon.schema.json)。コードで作業している場合Visual Studioこのファイルを使用して、JSON IntelliSenseを取得し、検証できます。</span><span class="sxs-lookup"><span data-stu-id="a3e82-149">If the link is not working, you can find the latest draft of the schema at [draft dynamic-ribbon.schema.json](https://github.com/OfficeDev/testing-assets/tree/master/jsonschema/dynamic-ribbon.schema.json).) If you are working in Visual Studio Code, you can use this file to get IntelliSense and to validate your JSON.</span></span> <span data-ttu-id="a3e82-150">詳細については、「コード - JSON スキーマと [設定を使用Visual Studio JSON](https://code.visualstudio.com/docs/languages/json#_json-schemas-and-settings)の編集」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="a3e82-150">For more information, see [Editing JSON with Visual Studio Code - JSON schemas and settings](https://code.visualstudio.com/docs/languages/json#_json-schemas-and-settings).</span></span>


1. <span data-ttu-id="a3e82-151">まず、次の 2 つの配列プロパティを持つ JSON 文字列を作成 `actions` します `tabs` 。</span><span class="sxs-lookup"><span data-stu-id="a3e82-151">Begin by creating a JSON string with two array properties named `actions` and `tabs`.</span></span> <span data-ttu-id="a3e82-152">配列 `actions` は、操作別タブのコントロールで実行できるすべての関数の仕様です。配列は、最大 10 までの 1 つ以上のコンテキスト タブ `tabs` *を定義します*。</span><span class="sxs-lookup"><span data-stu-id="a3e82-152">The `actions` array is a specification of all the functions that can be executed by controls on the contextual tab. The `tabs` array defines one or more contextual tabs, *up to a maximum of 10*.</span></span>

    ```json
    '{
      "actions": [

      ],
      "tabs": [

      ]
    }'
    ```

1. <span data-ttu-id="a3e82-153">この操作別タブの単純な例にはボタンが 1 つしか含めなく、したがってアクションは 1 つのみです。</span><span class="sxs-lookup"><span data-stu-id="a3e82-153">This simple example of a contextual tab will have only a single button and, thus, only a single action.</span></span> <span data-ttu-id="a3e82-154">以下を配列の唯一のメンバーとして追加 `actions` します。</span><span class="sxs-lookup"><span data-stu-id="a3e82-154">Add the following as the only member of the `actions` array.</span></span> <span data-ttu-id="a3e82-155">このマークアップについては、次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="a3e82-155">About this markup, note:</span></span>

    - <span data-ttu-id="a3e82-156">プロパティ `id` `type` とプロパティは必須です。</span><span class="sxs-lookup"><span data-stu-id="a3e82-156">The `id` and `type` properties are mandatory.</span></span>
    - <span data-ttu-id="a3e82-157">値には `type` 、"ExecuteFunction" または "ShowTaskpane" を指定できます。</span><span class="sxs-lookup"><span data-stu-id="a3e82-157">The value of `type` can be either "ExecuteFunction" or "ShowTaskpane".</span></span>
    - <span data-ttu-id="a3e82-158">プロパティ `functionName` は、値が次の場合にのみ使用 `type` されます `ExecuteFunction` 。</span><span class="sxs-lookup"><span data-stu-id="a3e82-158">The `functionName` property is only used when the value of `type` is `ExecuteFunction`.</span></span> <span data-ttu-id="a3e82-159">FunctionFile で定義されている関数の名前です。</span><span class="sxs-lookup"><span data-stu-id="a3e82-159">It is the name of a function defined in the FunctionFile.</span></span> <span data-ttu-id="a3e82-160">FunctionFile の詳細については、「アドイン コマンドの基本 [概念」を参照してください](add-in-commands.md)。</span><span class="sxs-lookup"><span data-stu-id="a3e82-160">For more information about the FunctionFile, see [Basic concepts for Add-in Commands](add-in-commands.md).</span></span>
    - <span data-ttu-id="a3e82-161">後の手順では、このアクションをコンテキスト タブのボタンにマップします。</span><span class="sxs-lookup"><span data-stu-id="a3e82-161">In a later step, you will map this action to a button on the contextual tab.</span></span>

    ```json
    {
      "id": "executeWriteData",
      "type": "ExecuteFunction",
      "functionName": "writeData"
    }
   ```

1. <span data-ttu-id="a3e82-162">以下を配列の唯一のメンバーとして追加 `tabs` します。</span><span class="sxs-lookup"><span data-stu-id="a3e82-162">Add the following as the only member of the `tabs` array.</span></span> <span data-ttu-id="a3e82-163">このマークアップについては、次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="a3e82-163">About this markup, note:</span></span>

    - <span data-ttu-id="a3e82-164">`id` プロパティは必須です。</span><span class="sxs-lookup"><span data-stu-id="a3e82-164">The `id` property is required.</span></span> <span data-ttu-id="a3e82-165">アドイン内のすべてのコンテキスト タブの中で一意である簡潔でわかりやすい ID を使用します。</span><span class="sxs-lookup"><span data-stu-id="a3e82-165">Use a brief, descriptive ID that is unique among all contextual tabs in your add-in.</span></span>
    - <span data-ttu-id="a3e82-166">`label` プロパティは必須です。</span><span class="sxs-lookup"><span data-stu-id="a3e82-166">The `label` property is required.</span></span> <span data-ttu-id="a3e82-167">コンテキスト タブのラベルとして使用すると、ユーザーに分け親しまれる文字列になります。</span><span class="sxs-lookup"><span data-stu-id="a3e82-167">It is a user-friendly string to serve as the label of the contextual tab.</span></span>
    - <span data-ttu-id="a3e82-168">`groups` プロパティは必須です。</span><span class="sxs-lookup"><span data-stu-id="a3e82-168">The `groups` property is required.</span></span> <span data-ttu-id="a3e82-169">タブに表示されるコントロールのグループを定義します。少なくとも 1 つのメンバーと *20 以下である必要があります*。</span><span class="sxs-lookup"><span data-stu-id="a3e82-169">It defines the groups of controls that will appear on the tab. It must have at least one member *and no more than 20*.</span></span> <span data-ttu-id="a3e82-170">(カスタム コンテキスト タブに設定できるコントロールの数にも制限があります。また、持っているグループの数も制限されます。</span><span class="sxs-lookup"><span data-stu-id="a3e82-170">(There are also limits on the number of controls that you can have on a custom contextual tab and that will also constrain how many groups that you have.</span></span> <span data-ttu-id="a3e82-171">詳細については、次の手順を参照してください)。</span><span class="sxs-lookup"><span data-stu-id="a3e82-171">See the next step for more information.)</span></span>

    > [!NOTE]
    > <span data-ttu-id="a3e82-172">タブ オブジェクトには、アドインの起動直後にタブを表示するかどうかを指定するオプションのプロパティ `visible` を指定することもできます。</span><span class="sxs-lookup"><span data-stu-id="a3e82-172">The tab object can also have an optional `visible` property that specifies whether the tab is visible immediately when the add-in starts up.</span></span> <span data-ttu-id="a3e82-173">コンテキスト タブは通常、ユーザー イベントが表示をトリガーするまで非表示になります (ユーザーがドキュメント内の何らかの種類のエンティティを選択した場合など)、プロパティは既定で存在しない場合に設定されます `visible` `false` 。</span><span class="sxs-lookup"><span data-stu-id="a3e82-173">Since contextual tabs are normally hidden until a user event triggers their visibility (such as the user selecting an entity of some type in the document), the `visible` property defaults to `false` when not present.</span></span> <span data-ttu-id="a3e82-174">後のセクションでは、イベントに応答してプロパティを `true` 設定する方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="a3e82-174">In a later section, we show how to set the property to `true` in response to an event.</span></span>

    ```json
    {
      "id": "CtxTab1",
      "label": "Data",
      "groups": [

      ]
    }
    ```

1. <span data-ttu-id="a3e82-175">単純な例では、コンテキスト タブには 1 つのグループのみがあります。</span><span class="sxs-lookup"><span data-stu-id="a3e82-175">In the simple ongoing example, the contextual tab has only a single group.</span></span> <span data-ttu-id="a3e82-176">以下を配列の唯一のメンバーとして追加 `groups` します。</span><span class="sxs-lookup"><span data-stu-id="a3e82-176">Add the following as the only member of the `groups` array.</span></span> <span data-ttu-id="a3e82-177">このマークアップについては、次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="a3e82-177">About this markup, note:</span></span>

    - <span data-ttu-id="a3e82-178">すべてのプロパティが必要です。</span><span class="sxs-lookup"><span data-stu-id="a3e82-178">All the properties are required.</span></span>
    - <span data-ttu-id="a3e82-179">この `id` プロパティは、タブ内のすべてのグループ間で一意である必要があります。簡潔でわかりやすい ID を使用します。</span><span class="sxs-lookup"><span data-stu-id="a3e82-179">The `id` property must be unique among all the groups in the tab. Use a brief, descriptive ID.</span></span>
    - <span data-ttu-id="a3e82-180">グループ `label` のラベルとして使用する、ユーザー に分かしい文字列です。</span><span class="sxs-lookup"><span data-stu-id="a3e82-180">The `label` is a user-friendly string to serve as the label of the group.</span></span>
    - <span data-ttu-id="a3e82-181">プロパティの値は、リボンのサイズとアプリケーション ウィンドウのサイズに応じてリボンに表示されるアイコンを指定するOffice `icon` 配列です。</span><span class="sxs-lookup"><span data-stu-id="a3e82-181">The `icon` property's value is an array of objects that specify the icons that the group will have on the ribbon depending on the size of the ribbon and the Office application window.</span></span>
    - <span data-ttu-id="a3e82-182">プロパティ `controls` の値は、グループ内のボタンとメニューを指定するオブジェクトの配列です。</span><span class="sxs-lookup"><span data-stu-id="a3e82-182">The `controls` property's value is an array of objects that specify the buttons and menus in the group.</span></span> <span data-ttu-id="a3e82-183">1 つのグループに少なくとも *1 つ、6 以下である必要があります*。</span><span class="sxs-lookup"><span data-stu-id="a3e82-183">There must be at least one and *no more than 6 in a group*.</span></span>

    > [!IMPORTANT]
    > <span data-ttu-id="a3e82-184">*タブ全体のコントロールの総数は 20 以下です。*</span><span class="sxs-lookup"><span data-stu-id="a3e82-184">*The total number of controls on the whole tab can be no more than 20.*</span></span> <span data-ttu-id="a3e82-185">たとえば、それぞれ 6 つのコントロールを持つ 3 つのグループと 2 つのコントロールを持つ 4 つ目のグループを持つ場合がありますが、4 つのグループにそれぞれ 6 つのコントロールを持つすることはできません。</span><span class="sxs-lookup"><span data-stu-id="a3e82-185">For example, you could have 3 groups with 6 controls each, and a fourth group with 2 controls, but you cannot have 4 groups with 6 controls each.</span></span>  

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

1. <span data-ttu-id="a3e82-186">すべてのグループには、32x32 px と 80x80 px の 2 つ以上のサイズのアイコンが必要です。</span><span class="sxs-lookup"><span data-stu-id="a3e82-186">Every group must have an icon of at least two sizes, 32x32 px and 80x80 px.</span></span> <span data-ttu-id="a3e82-187">必要に応じて、16x16 px、20x20 px、24x24 px、40x40 px、48x48 px、64x64 px のアイコンを設定できます。</span><span class="sxs-lookup"><span data-stu-id="a3e82-187">Optionally, you can also have icons of sizes 16x16 px, 20x20 px, 24x24 px, 40x40 px, 48x48 px, and 64x64 px.</span></span> <span data-ttu-id="a3e82-188">Office、リボンとアプリケーション ウィンドウのサイズに基づいて使用するアイコンOffice決定します。</span><span class="sxs-lookup"><span data-stu-id="a3e82-188">Office decides which icon to use based on the size of the ribbon and Office application window.</span></span> <span data-ttu-id="a3e82-189">アイコン配列に次のオブジェクトを追加します。</span><span class="sxs-lookup"><span data-stu-id="a3e82-189">Add the following objects to the icon array.</span></span> <span data-ttu-id="a3e82-190">(ウィンドウとリボンのサイズが、グループ上のコントロールの少なくとも1 つが表示されるのに十分な大きさの場合、グループ アイコンは表示されません。</span><span class="sxs-lookup"><span data-stu-id="a3e82-190">(If the window and ribbon sizes are large enough for at least one of the *controls* on the group to appear, then no group icon at all appears.</span></span> <span data-ttu-id="a3e82-191">たとえば、Word **ウィンドウを縮小** して展開する場合は、Word リボンの [スタイル] グループを参照してください)。このマークアップについては、次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="a3e82-191">For an example, watch the **Styles** group on the Word ribbon as you shrink and expand the Word window.) About this markup, note:</span></span>

    - <span data-ttu-id="a3e82-192">両方のプロパティが必要です。</span><span class="sxs-lookup"><span data-stu-id="a3e82-192">Both the properties are required.</span></span>
    - <span data-ttu-id="a3e82-193">プロパティ `size` の単位はピクセルです。</span><span class="sxs-lookup"><span data-stu-id="a3e82-193">The `size` property unit of measure is pixels.</span></span> <span data-ttu-id="a3e82-194">アイコンは常に正方形なので、数値は高さと幅の両方です。</span><span class="sxs-lookup"><span data-stu-id="a3e82-194">Icons are always square, so the number is both the height and the width.</span></span>
    - <span data-ttu-id="a3e82-195">この `sourceLocation` プロパティは、アイコンの完全な URL を指定します。</span><span class="sxs-lookup"><span data-stu-id="a3e82-195">The `sourceLocation` property specifies the full URL to the icon.</span></span>

    > [!IMPORTANT]
    > <span data-ttu-id="a3e82-196">開発から実稼働に移行する場合 (ドメインを localhost から contoso.com に変更する場合など) アドインのマニフェストの URL を通常は変更する必要があるのと同様に、コンテキスト タブ JSON の URL も変更する必要があります。</span><span class="sxs-lookup"><span data-stu-id="a3e82-196">Just as you typically must change the URLs in the add-in's manifest when you move from development to production (such as changing the domain from localhost to contoso.com), you must also change the URLs in your contextual tabs JSON.</span></span>

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

1. <span data-ttu-id="a3e82-197">この単純な例では、グループにボタンが 1 つしか表示されます。</span><span class="sxs-lookup"><span data-stu-id="a3e82-197">In our simple ongoing example, the group has only a single button.</span></span> <span data-ttu-id="a3e82-198">次のオブジェクトを配列の唯一のメンバーとして追加 `controls` します。</span><span class="sxs-lookup"><span data-stu-id="a3e82-198">Add the following object as the only member of the `controls` array.</span></span> <span data-ttu-id="a3e82-199">このマークアップについては、次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="a3e82-199">About this markup, note:</span></span>

    - <span data-ttu-id="a3e82-200">ただし、すべてのプロパティ `enabled` は必須です。</span><span class="sxs-lookup"><span data-stu-id="a3e82-200">All the properties, except `enabled`, are required.</span></span>
    - <span data-ttu-id="a3e82-201">`type` コントロールの種類を指定します。</span><span class="sxs-lookup"><span data-stu-id="a3e82-201">`type` specifies the type of control.</span></span> <span data-ttu-id="a3e82-202">値には、"Button"、"Menu"、または "MobileButton" を指定できます。</span><span class="sxs-lookup"><span data-stu-id="a3e82-202">The values can be "Button", "Menu", or "MobileButton".</span></span>
    - <span data-ttu-id="a3e82-203">`id` 最大 125 文字まで入力できます。</span><span class="sxs-lookup"><span data-stu-id="a3e82-203">`id` can be up to 125 characters.</span></span> 
    - <span data-ttu-id="a3e82-204">`actionId` は、配列で定義されているアクションの ID である必要 `actions` があります。</span><span class="sxs-lookup"><span data-stu-id="a3e82-204">`actionId` must be the ID of an action defined in the `actions` array.</span></span> <span data-ttu-id="a3e82-205">(このセクションの手順 1 を参照してください)。</span><span class="sxs-lookup"><span data-stu-id="a3e82-205">(See step 1 of this section.)</span></span>
    - <span data-ttu-id="a3e82-206">`label` は、ボタンのラベルとして使用する、ユーザー に使い分け可能な文字列です。</span><span class="sxs-lookup"><span data-stu-id="a3e82-206">`label` is a user-friendly string to serve as the label of the button.</span></span>
    - <span data-ttu-id="a3e82-207">`superTip` は、豊富な形式のツール ヒントを表します。</span><span class="sxs-lookup"><span data-stu-id="a3e82-207">`superTip` represents a rich form of tool tip.</span></span> <span data-ttu-id="a3e82-208">プロパティと `title` プロパティ `description` の両方が必要です。</span><span class="sxs-lookup"><span data-stu-id="a3e82-208">Both the `title` and `description` properties are required.</span></span>
    - <span data-ttu-id="a3e82-209">`icon` ボタンのアイコンを指定します。</span><span class="sxs-lookup"><span data-stu-id="a3e82-209">`icon` specifies the icons for the button.</span></span> <span data-ttu-id="a3e82-210">グループ アイコンに関する前の注釈もここに適用されます。</span><span class="sxs-lookup"><span data-stu-id="a3e82-210">The previous remarks about the group icon apply here too.</span></span>
    - <span data-ttu-id="a3e82-211">`enabled` (オプション) コンテキスト タブが表示される際にボタンを有効にするかどうかを指定します。</span><span class="sxs-lookup"><span data-stu-id="a3e82-211">`enabled` (optional) specifies whether the button is enabled when the contextual tab appears starts up.</span></span> <span data-ttu-id="a3e82-212">存在しない場合の既定値は次の値です `true` 。</span><span class="sxs-lookup"><span data-stu-id="a3e82-212">The default if not present is `true`.</span></span> 

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
 
<span data-ttu-id="a3e82-213">JSON BLOB の完全な例を次に示します。</span><span class="sxs-lookup"><span data-stu-id="a3e82-213">The following is the complete example of the JSON blob:</span></span>

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
      "label": "Data",
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

## <a name="register-the-contextual-tab-with-office-with-requestcreatecontrols"></a><span data-ttu-id="a3e82-214">requestCreateControls を使用してOfficeタブを登録する</span><span class="sxs-lookup"><span data-stu-id="a3e82-214">Register the contextual tab with Office with requestCreateControls</span></span>

<span data-ttu-id="a3e82-215">コンテキスト タブは [、Office.ribbon.requestCreateControls](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls_tabDefinition_) メソッドをOfficeして、コンテキスト タブに登録されます。</span><span class="sxs-lookup"><span data-stu-id="a3e82-215">The contextual tab is registered with Office by calling the [Office.ribbon.requestCreateControls](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls_tabDefinition_) method.</span></span> <span data-ttu-id="a3e82-216">これは通常、メソッドに割り当てられている関数またはメソッドで `Office.initialize` 行 `Office.onReady` われます。</span><span class="sxs-lookup"><span data-stu-id="a3e82-216">This is typically done in either the function that is assigned to `Office.initialize` or with the `Office.onReady` method.</span></span> <span data-ttu-id="a3e82-217">これらのメソッドとアドインの初期化の詳細については、「アドインの初期化Office [参照してください](../develop/initialize-add-in.md)。</span><span class="sxs-lookup"><span data-stu-id="a3e82-217">For more about these methods and initializing the add-in, see [Initialize your Office Add-in](../develop/initialize-add-in.md).</span></span> <span data-ttu-id="a3e82-218">ただし、初期化後はメソッドを呼び出す必要があります。</span><span class="sxs-lookup"><span data-stu-id="a3e82-218">You can, however, call the method anytime after initialization.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="a3e82-219">この `requestCreateControls` メソッドは、アドインの特定のセッションで 1 回だけ呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="a3e82-219">The `requestCreateControls` method can be called only once in a given session of an add-in.</span></span> <span data-ttu-id="a3e82-220">再度呼び出された場合は、エラーがスローされます。</span><span class="sxs-lookup"><span data-stu-id="a3e82-220">An error is thrown if it is called again.</span></span>

<span data-ttu-id="a3e82-221">次に例を示します。</span><span class="sxs-lookup"><span data-stu-id="a3e82-221">The following is an example.</span></span> <span data-ttu-id="a3e82-222">JSON 文字列を JavaScript 関数に渡す前に、メソッドを使用して `JSON.parse` JavaScript オブジェクトに変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="a3e82-222">Note that the JSON string must be converted to a JavaScript object with the `JSON.parse` method before it can be passed to a JavaScript function.</span></span>

```javascript
Office.onReady(async () => {
    const contextualTabJSON = ` ... `; // Assign the JSON string such as the one at the end of the preceding section.
    const contextualTab = JSON.parse(contextualTabJSON);
    await Office.ribbon.requestCreateControls(contextualTab);
});
```

## <a name="specify-the-contexts-when-the-tab-will-be-visible-with-requestupdate"></a><span data-ttu-id="a3e82-223">requestUpdate でタブが表示されるコンテキストを指定する</span><span class="sxs-lookup"><span data-stu-id="a3e82-223">Specify the contexts when the tab will be visible with requestUpdate</span></span>

<span data-ttu-id="a3e82-224">通常、カスタム コンテキスト タブは、ユーザーが開始するイベントによってアドインのコンテキストが変更されると表示されます。</span><span class="sxs-lookup"><span data-stu-id="a3e82-224">Typically, a custom contextual tab should appear when a user-initiated event changes the add-in context.</span></span> <span data-ttu-id="a3e82-225">(Excel ブックの既定のワークシートにある) グラフがアクティブ化されている場合にのみ、タブが表示されるシナリオを考えます。</span><span class="sxs-lookup"><span data-stu-id="a3e82-225">Consider a scenario in which the tab should be visible when, and only when, a chart (on the default worksheet of an Excel workbook) is activated.</span></span>

<span data-ttu-id="a3e82-226">まず、ハンドラーを割り当てる必要があります。</span><span class="sxs-lookup"><span data-stu-id="a3e82-226">Begin by assigning handlers.</span></span> <span data-ttu-id="a3e82-227">これは通常、次の例のようにメソッドで行われます。この例では、ハンドラー (後の手順で作成) をワークシート内のすべてのグラフのイベントに割り当 `Office.onReady` `onActivated` `onDeactivated` てる必要があります。</span><span class="sxs-lookup"><span data-stu-id="a3e82-227">This is commonly done in the `Office.onReady` method as in the following example which assigns handlers (created in a later step) to the `onActivated` and `onDeactivated` events of all the charts in the worksheet.</span></span>

```javascript
Office.onReady(async () => {
    const contextualTabJSON = ' ... '; // Assign the JSON string.
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

<span data-ttu-id="a3e82-228">次に、ハンドラーを定義します。</span><span class="sxs-lookup"><span data-stu-id="a3e82-228">Next, define the handlers.</span></span> <span data-ttu-id="a3e82-229">次に示すのは単純な例ですが、より堅牢なバージョンの関数については、この記事で後の `showDataTab` [「HostRestartNeeded](#handling-the-hostrestartneeded-error) エラーの処理」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="a3e82-229">The following is a simple example of a `showDataTab`, but see [Handling the HostRestartNeeded error](#handling-the-hostrestartneeded-error) later in this article for a more robust version of the function.</span></span> <span data-ttu-id="a3e82-230">このコードについては、以下の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="a3e82-230">About this code, note:</span></span>

- <span data-ttu-id="a3e82-231">Office では、リボンの状態を更新するタイミングが制御されます。</span><span class="sxs-lookup"><span data-stu-id="a3e82-231">Office controls when it updates the state of the ribbon.</span></span> <span data-ttu-id="a3e82-232">[Office.ribbon.requestUpdate](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestupdate-input-)メソッドは、更新要求をキューに入れられます。</span><span class="sxs-lookup"><span data-stu-id="a3e82-232">The  [Office.ribbon.requestUpdate](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestupdate-input-) method queues a request to update.</span></span> <span data-ttu-id="a3e82-233">このメソッドは、リボンが実際に更新されるのではなく、要求をキューに入れ次第、オブジェクト `Promise` を解決します。</span><span class="sxs-lookup"><span data-stu-id="a3e82-233">The method will resolve the `Promise` object as soon as it has queued the request, not when the ribbon actually updates.</span></span>
- <span data-ttu-id="a3e82-234">メソッドのパラメーターは `requestUpdate` [RibbonUpdaterData](/javascript/api/office/office.ribbonupdaterdata) オブジェクトで、(1) *JSON* で指定されているとおりに ID でタブを指定し、(2) タブの可視性を指定します。</span><span class="sxs-lookup"><span data-stu-id="a3e82-234">The parameter for the `requestUpdate` method is a [RibbonUpdaterData](/javascript/api/office/office.ribbonupdaterdata) object that (1) specifies the tab by its ID *exactly as specified in the JSON* and (2) specifies visibility of the tab.</span></span>
- <span data-ttu-id="a3e82-235">同じコンテキストで表示するカスタム コンテキスト タブが複数ある場合は、単純にタブ オブジェクトを配列に追加 `tabs` します。</span><span class="sxs-lookup"><span data-stu-id="a3e82-235">If you have more than one custom contextual tab that should be visible in the same context, you simply add additional tab objects to the `tabs` array.</span></span>

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

<span data-ttu-id="a3e82-236">タブを非表示にするハンドラーは、プロパティを設定し戻す以外は、ほぼ `visible` 同じです `false` 。</span><span class="sxs-lookup"><span data-stu-id="a3e82-236">The handler to hide the tab is nearly identical, except that it sets the `visible` property back to `false`.</span></span>

<span data-ttu-id="a3e82-237">またOffice JavaScript ライブラリには、オブジェクトの作成を容易にするためのインターフェイス (型) `RibbonUpdateData` がいくつか用意されています。</span><span class="sxs-lookup"><span data-stu-id="a3e82-237">The Office JavaScript library also provides several interfaces (types) to make it easier to construct the`RibbonUpdateData` object.</span></span> <span data-ttu-id="a3e82-238">TypeScript の `showDataTab` 関数を次に示します。この関数は、これらの型を利用します。</span><span class="sxs-lookup"><span data-stu-id="a3e82-238">The following is the `showDataTab` function in TypeScript and it makes use of these types.</span></span>

```typescript
const showDataTab = async () => {
    const myContextualTab: Tab = {id: "CtxTab1", visible: true};
    const ribbonUpdater: RibbonUpdaterData = { tabs: [ myContextualTab ]};
    await Office.ribbon.requestUpdate(ribbonUpdater);
}
```

### <a name="toggle-tab-visibility-and-the-enabled-status-of-a-button-at-the-same-time"></a><span data-ttu-id="a3e82-239">タブの表示とボタンの有効な状態を同時に切り替える</span><span class="sxs-lookup"><span data-stu-id="a3e82-239">Toggle tab visibility and the enabled status of a button at the same time</span></span>

<span data-ttu-id="a3e82-240">このメソッドは、カスタム コンテキスト タブまたはカスタム コア タブのカスタム ボタンの有効または無効の状態を切り替 `requestUpdate` える場合にも使用されます。詳細については、「アドイン コマンドを [有効または無効にする」を参照してください](disable-add-in-commands.md)。</span><span class="sxs-lookup"><span data-stu-id="a3e82-240">The `requestUpdate` method is also used to toggle the enabled or disabled status of a custom button on either a custom contextual tab or a custom core tab. For details about this, see [Enable and Disable Add-in Commands](disable-add-in-commands.md).</span></span> <span data-ttu-id="a3e82-241">タブの表示とボタンの有効な状態の両方を同時に変更するシナリオが考えられます。</span><span class="sxs-lookup"><span data-stu-id="a3e82-241">There may be scenarios in which you want to change both the visibility of a tab and the enabled status of a button at the same time.</span></span> <span data-ttu-id="a3e82-242">これは、1 回の呼び出しで行います `requestUpdate` 。</span><span class="sxs-lookup"><span data-stu-id="a3e82-242">You can do this with a single call of `requestUpdate`.</span></span> <span data-ttu-id="a3e82-243">次の例では、コンテキスト タブが表示されるのと同時に、コア タブのボタンが有効になります。</span><span class="sxs-lookup"><span data-stu-id="a3e82-243">The following is an example in which a button on a core tab is enabled at the same time as a contextual tab is made visible.</span></span>

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
                controls: [
                {
                    id: "MyButton",
                    enabled: true
                }
            ]}
        ]});
}
```

<span data-ttu-id="a3e82-244">次の例では、有効になっているボタンは、表示されているのと同じコンテキスト タブにあります。</span><span class="sxs-lookup"><span data-stu-id="a3e82-244">In the following example, the button that is enabled is on the very same contextual tab that is being made visible.</span></span>

```javascript
function myContextChanges() {
    Office.ribbon.requestUpdate({
        tabs: [
            {
                id: "CtxTab1",
                visible: true,
                controls: [
                    {
                        id: "MyButton",
                        enabled: true
                    }
                ]
            }
        ]});
}
```

## <a name="localizing-the-json-blob"></a><span data-ttu-id="a3e82-245">JSON BLOB のローカライズ</span><span class="sxs-lookup"><span data-stu-id="a3e82-245">Localizing the JSON blob</span></span>

<span data-ttu-id="a3e82-246">渡される JSON BLOB は、カスタム コア タブのマニフェスト マークアップがローカライズされるのと同じ方法ではローカライズされません (マニフェストからのコントロールのローカライズで `requestCreateControls` [説明します](../develop/localization.md#control-localization-from-the-manifest))。</span><span class="sxs-lookup"><span data-stu-id="a3e82-246">The JSON blob that is passed to `requestCreateControls` is not localized the same way that the manifest markup for custom core tabs is localized (which is described at [Control localization from the manifest](../develop/localization.md#control-localization-from-the-manifest)).</span></span> <span data-ttu-id="a3e82-247">代わりに、ローカライズは、ロケールごとに異なる JSON BLOB を使用して実行時に行う必要があります。</span><span class="sxs-lookup"><span data-stu-id="a3e82-247">Instead, the localization must occur at runtime using distinct JSON blobs for each locale.</span></span> <span data-ttu-id="a3e82-248">`switch` [Office.context.displayLanguage](/javascript/api/office/office.context#displayLanguage)プロパティをテストするステートメントを使用してください。</span><span class="sxs-lookup"><span data-stu-id="a3e82-248">We suggest that you use a `switch` statement that tests the [Office.context.displayLanguage](/javascript/api/office/office.context#displayLanguage) property.</span></span> <span data-ttu-id="a3e82-249">例を次に示します。</span><span class="sxs-lookup"><span data-stu-id="a3e82-249">The following is an example:</span></span>

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
                          "label": "Data",
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
                          "label": "Données",
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

<span data-ttu-id="a3e82-250">次に、次の例のように、コードで関数を呼び出して、渡されるローカライズされた BLOB `requestCreateControls` を取得します。</span><span class="sxs-lookup"><span data-stu-id="a3e82-250">Then your code calls the function to get the localized blob that is passed to `requestCreateControls`, as in the following example:</span></span>

```javascript
var contextualTabJSON = GetContextualTabsJsonSupportedLocale();
```

## <a name="handling-the-hostrestartneeded-error"></a><span data-ttu-id="a3e82-251">HostRestartNeeded エラーの処理</span><span class="sxs-lookup"><span data-stu-id="a3e82-251">Handling the HostRestartNeeded error</span></span>

<span data-ttu-id="a3e82-252">一部のシナリオでは、Office はリボンを更新できず、エラーを返します。</span><span class="sxs-lookup"><span data-stu-id="a3e82-252">In some scenarios, Office is unable to update the ribbon and will return an error.</span></span> <span data-ttu-id="a3e82-253">たとえば、アドインがアップグレードされ、アップグレードされたアドインに異なるカスタム アドイン コマンドのセットがある場合は、Office アプリケーションを閉じてから、もう一度開く必要があります。</span><span class="sxs-lookup"><span data-stu-id="a3e82-253">For example, if the add-in is upgraded and the upgraded add-in has a different set of custom add-in commands, then the Office application must be closed and reopened.</span></span> <span data-ttu-id="a3e82-254">それまでの間、`requestUpdate` メソッドは `HostRestartNeeded` エラーを返します。</span><span class="sxs-lookup"><span data-stu-id="a3e82-254">Until it is, the `requestUpdate` method will return the error `HostRestartNeeded`.</span></span> <span data-ttu-id="a3e82-255">このエラーの処理方法の例を次に示します。</span><span class="sxs-lookup"><span data-stu-id="a3e82-255">The following is an example of how to handle this error.</span></span> <span data-ttu-id="a3e82-256">この場合、`reportError` メソッドがユーザーにエラーを表示します。</span><span class="sxs-lookup"><span data-stu-id="a3e82-256">In this case, the `reportError` method displays the error to the user.</span></span>

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
