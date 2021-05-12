---
title: カスタム コンテキスト タブを Officeアドインで作成する
description: カスタム コンテキスト タブをアドインに追加するOffice説明します。
ms.date: 05/12/2021
localization_priority: Normal
ms.openlocfilehash: 980beb24a3d384ecf21da44db288272a1ab1b0e3
ms.sourcegitcommit: 30f6c620380075e3459cac748ca0c656427b384d
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/12/2021
ms.locfileid: "52330172"
---
# <a name="create-custom-contextual-tabs-in-office-add-ins"></a><span data-ttu-id="ab706-103">カスタム コンテキスト タブを Officeアドインで作成する</span><span class="sxs-lookup"><span data-stu-id="ab706-103">Create custom contextual tabs in Office Add-ins</span></span>

<span data-ttu-id="ab706-104">コンテキスト タブは、指定したイベントがドキュメント内で発生した場合にタブ行に表示Officeリボンの非表示のタブ コントロールOfficeです。</span><span class="sxs-lookup"><span data-stu-id="ab706-104">A contextual tab is a hidden tab control in the Office ribbon that is displayed in the tab row when a specified event occurs in the Office document.</span></span> <span data-ttu-id="ab706-105">たとえば、テーブル **が選択されている** ときにリボンのExcel[テーブルのデザイン] タブが表示されます。</span><span class="sxs-lookup"><span data-stu-id="ab706-105">For example, the **Table Design** tab that appears on the Excel ribbon when a table is selected.</span></span> <span data-ttu-id="ab706-106">表示を変更するイベント ハンドラーを作成することで、Office アドインにカスタム コンテキスト タブを含め、表示または非表示の設定を指定できます。</span><span class="sxs-lookup"><span data-stu-id="ab706-106">You can include custom contextual tabs in your Office Add-in and specify when they are visible or hidden, by creating event handlers that change the visibility.</span></span> <span data-ttu-id="ab706-107">(ただし、カスタム コンテキスト タブはフォーカスの変更に応答しない)。</span><span class="sxs-lookup"><span data-stu-id="ab706-107">(However, custom contextual tabs do not respond to focus changes.)</span></span>

> [!NOTE]
> <span data-ttu-id="ab706-108">この記事は、以下のドキュメントについて既に理解していることを前提としています。</span><span class="sxs-lookup"><span data-stu-id="ab706-108">This article assumes that you are familiar with the following documentation.</span></span> <span data-ttu-id="ab706-109">最近、アドイン コマンド (カスタム メニュー項目とリボン ボタン) を使用してない場合は、ドキュメントをご確認ください。</span><span class="sxs-lookup"><span data-stu-id="ab706-109">Please review it if you haven't worked with Add-in Commands (custom menu items and ribbon buttons) recently.</span></span>
>
> - [<span data-ttu-id="ab706-110">アドイン コマンドの基本概念</span><span class="sxs-lookup"><span data-stu-id="ab706-110">Basic concepts for Add-in Commands</span></span>](add-in-commands.md)

> [!IMPORTANT]
> <span data-ttu-id="ab706-111">カスタム コンテキスト タブは現在、次のプラットフォームExcelビルドでのみサポートされています。</span><span class="sxs-lookup"><span data-stu-id="ab706-111">Custom contextual tabs are currently only supported on Excel and only on these platforms and builds:</span></span>
>
> - <span data-ttu-id="ab706-112">Excel (Windows サブスクリプションMicrosoft 365): バージョン 2102 (ビルド 13801.20294) 以降。</span><span class="sxs-lookup"><span data-stu-id="ab706-112">Excel on Windows (Microsoft 365 subscription only): Version 2102 (Build 13801.20294) or later.</span></span>
> - <span data-ttu-id="ab706-113">Excel on the web</span><span class="sxs-lookup"><span data-stu-id="ab706-113">Excel on the web</span></span>

> [!NOTE]
> <span data-ttu-id="ab706-114">カスタム コンテキスト タブは、次の要件セットをサポートするプラットフォームでのみ機能します。</span><span class="sxs-lookup"><span data-stu-id="ab706-114">Custom contextual tabs work only on platforms that support the following requirement sets.</span></span> <span data-ttu-id="ab706-115">要件セットとそれらを使用する方法の詳細については、「アプリケーションと API の要件Office[を指定する」を参照してください](../develop/specify-office-hosts-and-api-requirements.md)。</span><span class="sxs-lookup"><span data-stu-id="ab706-115">For more about requirement sets and how to work with them, see [Specify Office applications and API requirements](../develop/specify-office-hosts-and-api-requirements.md).</span></span>
>
> - [<span data-ttu-id="ab706-116">RibbonApi 1.2</span><span class="sxs-lookup"><span data-stu-id="ab706-116">RibbonApi 1.2</span></span>](../reference/requirement-sets/ribbon-api-requirement-sets.md)
> - [<span data-ttu-id="ab706-117">SharedRuntime 1.1</span><span class="sxs-lookup"><span data-stu-id="ab706-117">SharedRuntime 1.1</span></span>](../reference/requirement-sets/shared-runtime-requirement-sets.md)
>
> <span data-ttu-id="ab706-118">コードのランタイム チェックを使用して、ユーザーのホストとプラットフォームの組み合わせがこれらの要件セットをサポートするかどうかをテストできます(「アプリケーションと[API](../develop/specify-office-hosts-and-api-requirements.md#use-runtime-checks-in-your-javascript-code)の要件を指定Officeを指定する」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="ab706-118">You can use the runtime checks in your code to test whether the user's host and platform combination supports these requirement sets as described in [Specify Office applications and API requirements](../develop/specify-office-hosts-and-api-requirements.md#use-runtime-checks-in-your-javascript-code).</span></span> <span data-ttu-id="ab706-119">(マニフェストで要件セットを指定する方法は、この記事でも説明しますが、現在 RibbonApi 1.2 では機能しません)。または、カスタム コンテキスト タブがサポートされていない場合に、別の [UI エクスペリエンスを実装することもできます](#implement-an-alternate-ui-experience-when-custom-contextual-tabs-are-not-supported)。</span><span class="sxs-lookup"><span data-stu-id="ab706-119">(The technique of specifying the requirement sets in the manifest, which is also described in that article, does not currently work for RibbonApi 1.2.) Alternatively, you can [implement an alternate UI experience when custom contextual tabs are not supported](#implement-an-alternate-ui-experience-when-custom-contextual-tabs-are-not-supported).</span></span>

## <a name="behavior-of-custom-contextual-tabs"></a><span data-ttu-id="ab706-120">カスタム コンテキスト タブの動作</span><span class="sxs-lookup"><span data-stu-id="ab706-120">Behavior of custom contextual tabs</span></span>

<span data-ttu-id="ab706-121">カスタム コンテキスト タブのユーザー エクスペリエンスは、組み込みのコンテキスト タブのOfficeに従います。</span><span class="sxs-lookup"><span data-stu-id="ab706-121">The user experience for custom contextual tabs follows the pattern of built-in Office contextual tabs.</span></span> <span data-ttu-id="ab706-122">配置のカスタム コンテキスト タブの基本的な原則を次に示します。</span><span class="sxs-lookup"><span data-stu-id="ab706-122">The following are the basic principles for the placement custom contextual tabs:</span></span>

- <span data-ttu-id="ab706-123">カスタム コンテキスト タブが表示されている場合は、リボンの右側に表示されます。</span><span class="sxs-lookup"><span data-stu-id="ab706-123">When a custom contextual tab is visible, it appears on the right end of the ribbon.</span></span>
- <span data-ttu-id="ab706-124">1 つ以上の組み込みのコンテキスト タブと、アドインから 1 つ以上のカスタム コンテキスト タブが同時に表示される場合、カスタム コンテキスト タブは常にすべての組み込みコンテキスト タブの右側に表示されます。</span><span class="sxs-lookup"><span data-stu-id="ab706-124">If one or more built-in contextual tabs and one or more custom contextual tabs from add-ins are visible at the same time, the custom contextual tabs are always to the right of all of the built-in contextual tabs.</span></span>
- <span data-ttu-id="ab706-125">アドインに複数のコンテキスト タブが含み、複数のコンテキストが表示されている場合は、アドインで定義されている順序で表示されます。</span><span class="sxs-lookup"><span data-stu-id="ab706-125">If your add-in has more than one contextual tab and there are contexts in which more than one is visible, they appear in the order in which they are defined in your add-in.</span></span> <span data-ttu-id="ab706-126">(方向は Office 言語と同じ方向です。つまり、左から右の言語では左から右、右から左の言語では右から左)。定義[方法の詳細については、「タブに表示される](#define-the-groups-and-controls-that-appear-on-the-tab)グループとコントロールを定義する」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="ab706-126">(The direction is the same direction as the Office language; that is, is left-to-right in left-to-right languages, but right-to-left in right-to-left languages.) See [Define the groups and controls that appear on the tab](#define-the-groups-and-controls-that-appear-on-the-tab) for details about how you define them.</span></span>
- <span data-ttu-id="ab706-127">特定のコンテキストで表示されるコンテキスト タブが複数のアドインにある場合は、アドインが起動された順序で表示されます。</span><span class="sxs-lookup"><span data-stu-id="ab706-127">If more than one add-in has a contextual tab that is visible in a specific context, then they appear in the order in which the add-ins were launched.</span></span>
- <span data-ttu-id="ab706-128">カスタム *コンテキスト* タブは、カスタム コア タブとは異なり、アプリケーションのリボンOffice完全には追加されません。</span><span class="sxs-lookup"><span data-stu-id="ab706-128">Custom *contextual* tabs, unlike custom core tabs, are not added permanently to the Office application's ribbon.</span></span> <span data-ttu-id="ab706-129">アドインが実行されているOfficeドキュメントにのみ存在します。</span><span class="sxs-lookup"><span data-stu-id="ab706-129">They are present only in Office documents on which your add-in is running.</span></span>

## <a name="major-steps-for-including-a-contextual-tab-in-an-add-in"></a><span data-ttu-id="ab706-130">アドインにコンテキスト タブを含む場合の主な手順</span><span class="sxs-lookup"><span data-stu-id="ab706-130">Major steps for including a contextual tab in an add-in</span></span>

<span data-ttu-id="ab706-131">アドインにカスタム コンテキスト タブを含む場合の主な手順を次に示します。</span><span class="sxs-lookup"><span data-stu-id="ab706-131">The following are the major steps for including a custom contextual tab in an add-in:</span></span>

1. <span data-ttu-id="ab706-132">共有ランタイムを使用するアドインを構成します。</span><span class="sxs-lookup"><span data-stu-id="ab706-132">Configure the add-in to use a shared runtime.</span></span>
1. <span data-ttu-id="ab706-133">タブと、タブに表示されるグループとコントロールを定義します。</span><span class="sxs-lookup"><span data-stu-id="ab706-133">Define the tab and the groups and controls that appear on it.</span></span>
1. <span data-ttu-id="ab706-134">コンテキスト タブを [コンテキスト] タブにOffice。</span><span class="sxs-lookup"><span data-stu-id="ab706-134">Register the contextual tab with Office.</span></span>
1. <span data-ttu-id="ab706-135">タブが表示される状況を指定します。</span><span class="sxs-lookup"><span data-stu-id="ab706-135">Specify the circumstances when the tab will be visible.</span></span>

## <a name="configure-the-add-in-to-use-a-shared-runtime"></a><span data-ttu-id="ab706-136">共有ランタイムを使用するアドインを構成する</span><span class="sxs-lookup"><span data-stu-id="ab706-136">Configure the add-in to use a shared runtime</span></span>

<span data-ttu-id="ab706-137">カスタム コンテキスト タブを追加するには、共有ランタイムを使用するアドインが必要です。</span><span class="sxs-lookup"><span data-stu-id="ab706-137">Adding custom contextual tabs requires your add-in to use the shared runtime.</span></span> <span data-ttu-id="ab706-138">詳細については、「共有ランタイム [を使用するアドインを構成する」を参照してください](../develop/configure-your-add-in-to-use-a-shared-runtime.md)。</span><span class="sxs-lookup"><span data-stu-id="ab706-138">For more information, see [Configure an add-in to use a shared runtime](../develop/configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

## <a name="define-the-groups-and-controls-that-appear-on-the-tab"></a><span data-ttu-id="ab706-139">タブに表示されるグループとコントロールを定義する</span><span class="sxs-lookup"><span data-stu-id="ab706-139">Define the groups and controls that appear on the tab</span></span>

<span data-ttu-id="ab706-140">マニフェストで XML で定義されたカスタム コア タブとは異なり、カスタム コンテキスト タブは実行時に JSON BLOB を使用して定義されます。</span><span class="sxs-lookup"><span data-stu-id="ab706-140">Unlike custom core tabs, which are defined with XML in the manifest, custom contextual tabs are defined at runtime with a JSON blob.</span></span> <span data-ttu-id="ab706-141">コードは BLOB を JavaScript オブジェクトに解析し、オブジェクトを[Office.ribbon.requestCreateControls メソッドに渡](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls-tabDefinition-)します。</span><span class="sxs-lookup"><span data-stu-id="ab706-141">Your code parses the blob into a JavaScript object, and then passes the object to the [Office.ribbon.requestCreateControls](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls-tabDefinition-) method.</span></span> <span data-ttu-id="ab706-142">カスタム コンテキスト タブは、アドインが現在実行されているドキュメントにのみ表示されます。</span><span class="sxs-lookup"><span data-stu-id="ab706-142">Custom contextual tabs are only present in documents on which your add-in is currently running.</span></span> <span data-ttu-id="ab706-143">これは、アドインのインストール時に Office アプリケーション リボンに追加されるカスタム コア タブとは異なります。また、別のドキュメントを開いた時点でも存在します。</span><span class="sxs-lookup"><span data-stu-id="ab706-143">This is different from custom core tabs which are added to the Office application ribbon when the add-in is installed and remain present when another document is opened.</span></span> <span data-ttu-id="ab706-144">また、 `requestCreateControls` メソッドはアドインのセッションで 1 回だけ実行できます。</span><span class="sxs-lookup"><span data-stu-id="ab706-144">Also, the `requestCreateControls` method can be run only once in a session of your add-in.</span></span> <span data-ttu-id="ab706-145">再び呼び出された場合は、エラーがスローされます。</span><span class="sxs-lookup"><span data-stu-id="ab706-145">If it is called again, an error is thrown.</span></span>

> [!NOTE]
> <span data-ttu-id="ab706-146">JSON BLOB のプロパティとサブプロパティ (およびキー名) の構造は、マニフェスト XML の [CustomTab](../reference/manifest/customtab.md) 要素とその子孫要素の構造と大まかに平行です。</span><span class="sxs-lookup"><span data-stu-id="ab706-146">The structure of the JSON blob's properties and subproperties (and the key names) is roughly parallel to the structure of the [CustomTab](../reference/manifest/customtab.md) element and its descendant elements in the manifest XML.</span></span>

<span data-ttu-id="ab706-147">コンテキスト タブ JSON BLOB のステップ バイ ステップの例を作成します。</span><span class="sxs-lookup"><span data-stu-id="ab706-147">We'll construct an example of a contextual tabs JSON blob step-by-step.</span></span> <span data-ttu-id="ab706-148">(コンテキスト タブ JSON の完全なスキーマは、dynamic-ribbon.schema.js[ です](https://developer.microsoft.com/json-schemas/office-js/dynamic-ribbon.schema.json)。</span><span class="sxs-lookup"><span data-stu-id="ab706-148">(The full schema for the contextual tab JSON is at [dynamic-ribbon.schema.json](https://developer.microsoft.com/json-schemas/office-js/dynamic-ribbon.schema.json).</span></span> <span data-ttu-id="ab706-149">このリンクは、コンテキスト タブのプレビュー期間中に機能していない可能性があります。</span><span class="sxs-lookup"><span data-stu-id="ab706-149">This link may not be working in the preview period for contextual tabs.</span></span> <span data-ttu-id="ab706-150">リンクが機能していない場合は、下書きでスキーマの最新の下書きを見dynamic-ribbon.schema.js[を参照](https://github.com/OfficeDev/testing-assets/tree/master/jsonschema/dynamic-ribbon/1.0/dynamic-ribbon.schema.json)してください。)このドキュメントで作業しているVisual Studio Code、このファイルを使用して JSON のIntelliSense検証できます。</span><span class="sxs-lookup"><span data-stu-id="ab706-150">If the link is not working, you can find the latest draft of the schema at [draft dynamic-ribbon.schema.json](https://github.com/OfficeDev/testing-assets/tree/master/jsonschema/dynamic-ribbon/1.0/dynamic-ribbon.schema.json).) If you are working in Visual Studio Code, you can use this file to get IntelliSense and to validate your JSON.</span></span> <span data-ttu-id="ab706-151">詳細については、「JSON スキーマと設定を使用Visual Studio Code JSON の編集[」を参照してください](https://code.visualstudio.com/docs/languages/json#_json-schemas-and-settings)。</span><span class="sxs-lookup"><span data-stu-id="ab706-151">For more information, see [Editing JSON with Visual Studio Code - JSON schemas and settings](https://code.visualstudio.com/docs/languages/json#_json-schemas-and-settings).</span></span>


1. <span data-ttu-id="ab706-152">まず、という名前の 2 つの配列プロパティを持つ JSON 文字列を作成 `actions` します `tabs` 。</span><span class="sxs-lookup"><span data-stu-id="ab706-152">Begin by creating a JSON string with two array properties named `actions` and `tabs`.</span></span> <span data-ttu-id="ab706-153">配列 `actions` は、コンテキスト タブのコントロールで実行できるすべての関数の仕様です。配列 `tabs` は、最大 *20* までの 1 つ以上のコンテキスト タブを定義します。</span><span class="sxs-lookup"><span data-stu-id="ab706-153">The `actions` array is a specification of all the functions that can be executed by controls on the contextual tab. The `tabs` array defines one or more contextual tabs, *up to a maximum of 20*.</span></span>

    ```json
    '{
      "actions": [

      ],
      "tabs": [

      ]
    }'
    ```

1. <span data-ttu-id="ab706-154">コンテキスト タブのこの簡単な例では、ボタンが 1 つしか表示され、1 つのアクションだけが表示されます。</span><span class="sxs-lookup"><span data-stu-id="ab706-154">This simple example of a contextual tab will have only a single button and, thus, only a single action.</span></span> <span data-ttu-id="ab706-155">配列の唯一のメンバーとして次を追加 `actions` します。</span><span class="sxs-lookup"><span data-stu-id="ab706-155">Add the following as the only member of the `actions` array.</span></span> <span data-ttu-id="ab706-156">このマークアップについて、次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="ab706-156">About this markup, note:</span></span>

    - <span data-ttu-id="ab706-157">プロパティ `id` と `type` プロパティは必須です。</span><span class="sxs-lookup"><span data-stu-id="ab706-157">The `id` and `type` properties are mandatory.</span></span>
    - <span data-ttu-id="ab706-158">値には `type` 、"ExecuteFunction" または "ShowTaskpane" を指定できます。</span><span class="sxs-lookup"><span data-stu-id="ab706-158">The value of `type` can be either "ExecuteFunction" or "ShowTaskpane".</span></span>
    - <span data-ttu-id="ab706-159">プロパティ `functionName` は、 の値が . の場合にのみ `type` 使用されます `ExecuteFunction` 。</span><span class="sxs-lookup"><span data-stu-id="ab706-159">The `functionName` property is only used when the value of `type` is `ExecuteFunction`.</span></span> <span data-ttu-id="ab706-160">これは、FunctionFile で定義されている関数の名前です。</span><span class="sxs-lookup"><span data-stu-id="ab706-160">It is the name of a function defined in the FunctionFile.</span></span> <span data-ttu-id="ab706-161">FunctionFile の詳細については、「アドイン コマンドの基本的 [な概念」を参照してください](add-in-commands.md)。</span><span class="sxs-lookup"><span data-stu-id="ab706-161">For more information about the FunctionFile, see [Basic concepts for Add-in Commands](add-in-commands.md).</span></span>
    - <span data-ttu-id="ab706-162">後の手順では、このアクションをコンテキスト タブのボタンにマップします。</span><span class="sxs-lookup"><span data-stu-id="ab706-162">In a later step, you will map this action to a button on the contextual tab.</span></span>

    ```json
    {
      "id": "executeWriteData",
      "type": "ExecuteFunction",
      "functionName": "writeData"
    }
   ```

1. <span data-ttu-id="ab706-163">配列の唯一のメンバーとして次を追加 `tabs` します。</span><span class="sxs-lookup"><span data-stu-id="ab706-163">Add the following as the only member of the `tabs` array.</span></span> <span data-ttu-id="ab706-164">このマークアップについて、次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="ab706-164">About this markup, note:</span></span>

    - <span data-ttu-id="ab706-165">`id` プロパティは必須です。</span><span class="sxs-lookup"><span data-stu-id="ab706-165">The `id` property is required.</span></span> <span data-ttu-id="ab706-166">アドイン内のすべてのコンテキスト タブで一意の簡潔でわかりやすい ID を使用します。</span><span class="sxs-lookup"><span data-stu-id="ab706-166">Use a brief, descriptive ID that is unique among all contextual tabs in your add-in.</span></span>
    - <span data-ttu-id="ab706-167">`label` プロパティは必須です。</span><span class="sxs-lookup"><span data-stu-id="ab706-167">The `label` property is required.</span></span> <span data-ttu-id="ab706-168">コンテキスト タブのラベルとして機能するユーザーフレンドリーな文字列です。</span><span class="sxs-lookup"><span data-stu-id="ab706-168">It is a user-friendly string to serve as the label of the contextual tab.</span></span>
    - <span data-ttu-id="ab706-169">`groups` プロパティは必須です。</span><span class="sxs-lookup"><span data-stu-id="ab706-169">The `groups` property is required.</span></span> <span data-ttu-id="ab706-170">タブに表示されるコントロールのグループを定義します。少なくとも 1 つのメンバーと *20 以下のメンバーが必要です*。</span><span class="sxs-lookup"><span data-stu-id="ab706-170">It defines the groups of controls that will appear on the tab. It must have at least one member *and no more than 20*.</span></span> <span data-ttu-id="ab706-171">(カスタム コンテキスト タブで使用できるコントロールの数にも制限があります。また、ユーザーが持つグループの数も制限されます。</span><span class="sxs-lookup"><span data-stu-id="ab706-171">(There are also limits on the number of controls that you can have on a custom contextual tab and that will also constrain how many groups that you have.</span></span> <span data-ttu-id="ab706-172">詳細については、次の手順を参照してください)。</span><span class="sxs-lookup"><span data-stu-id="ab706-172">See the next step for more information.)</span></span>

    > [!NOTE]
    > <span data-ttu-id="ab706-173">Tab オブジェクトには、アドインの起動時にタブをすぐに表示するかどうかを指定するオプション `visible` のプロパティを指定することもできます。</span><span class="sxs-lookup"><span data-stu-id="ab706-173">The tab object can also have an optional `visible` property that specifies whether the tab is visible immediately when the add-in starts up.</span></span> <span data-ttu-id="ab706-174">コンテキスト タブは通常、ユーザー イベントが表示をトリガーするまで非表示になります (ドキュメント内で何らかの種類のエンティティを選択するユーザーなど)、プロパティの既定値は存在しない場合です `visible` `false` 。</span><span class="sxs-lookup"><span data-stu-id="ab706-174">Since contextual tabs are normally hidden until a user event triggers their visibility (such as the user selecting an entity of some type in the document), the `visible` property defaults to `false` when not present.</span></span> <span data-ttu-id="ab706-175">後のセクションでは、イベントに応答してプロパティを設定 `true` する方法を示します。</span><span class="sxs-lookup"><span data-stu-id="ab706-175">In a later section, we show how to set the property to `true` in response to an event.</span></span>

    ```json
    {
      "id": "CtxTab1",
      "label": "Contoso Data",
      "groups": [

      ]
    }
    ```

1. <span data-ttu-id="ab706-176">単純な進行中の例では、コンテキスト タブには 1 つのグループのみがあります。</span><span class="sxs-lookup"><span data-stu-id="ab706-176">In the simple ongoing example, the contextual tab has only a single group.</span></span> <span data-ttu-id="ab706-177">配列の唯一のメンバーとして次を追加 `groups` します。</span><span class="sxs-lookup"><span data-stu-id="ab706-177">Add the following as the only member of the `groups` array.</span></span> <span data-ttu-id="ab706-178">このマークアップについて、次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="ab706-178">About this markup, note:</span></span>

    - <span data-ttu-id="ab706-179">すべてのプロパティが必要です。</span><span class="sxs-lookup"><span data-stu-id="ab706-179">All the properties are required.</span></span>
    - <span data-ttu-id="ab706-180">プロパティ `id` は、タブ内のすべてのグループ間で一意である必要があります。簡潔でわかりやすい ID を使用します。</span><span class="sxs-lookup"><span data-stu-id="ab706-180">The `id` property must be unique among all the groups in the tab. Use a brief, descriptive ID.</span></span>
    - <span data-ttu-id="ab706-181">グループ `label` のラベルとして使用するユーザーフレンドリーな文字列です。</span><span class="sxs-lookup"><span data-stu-id="ab706-181">The `label` is a user-friendly string to serve as the label of the group.</span></span>
    - <span data-ttu-id="ab706-182">プロパティの値は、リボンのサイズとアプリケーション ウィンドウのサイズに応じて、グループがリボンに表示するアイコンをOffice `icon` 配列です。</span><span class="sxs-lookup"><span data-stu-id="ab706-182">The `icon` property's value is an array of objects that specify the icons that the group will have on the ribbon depending on the size of the ribbon and the Office application window.</span></span>
    - <span data-ttu-id="ab706-183">プロパティ `controls` の値は、グループ内のボタンとメニューを指定するオブジェクトの配列です。</span><span class="sxs-lookup"><span data-stu-id="ab706-183">The `controls` property's value is an array of objects that specify the buttons and menus in the group.</span></span> <span data-ttu-id="ab706-184">少なくとも 1 つが必要です。</span><span class="sxs-lookup"><span data-stu-id="ab706-184">There must be at least one.</span></span>

    > [!IMPORTANT]
    > <span data-ttu-id="ab706-185">*タブ全体のコントロールの総数は 20 以下です。*</span><span class="sxs-lookup"><span data-stu-id="ab706-185">*The total number of controls on the whole tab can be no more than 20.*</span></span> <span data-ttu-id="ab706-186">たとえば、各コントロールが 6 つのグループが 3 つ、コントロールが 2 つの 4 番目のグループを持つグループを 3 つ持つ場合がありますが、6 つのコントロールを持つグループを 4 つ持つ必要があります。</span><span class="sxs-lookup"><span data-stu-id="ab706-186">For example, you could have 3 groups with 6 controls each, and a fourth group with 2 controls, but you cannot have 4 groups with 6 controls each.</span></span>  

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

1. <span data-ttu-id="ab706-187">すべてのグループには、32x32 px と 80x80 px の少なくとも 2 つのサイズのアイコンが必要です。</span><span class="sxs-lookup"><span data-stu-id="ab706-187">Every group must have an icon of at least two sizes, 32x32 px and 80x80 px.</span></span> <span data-ttu-id="ab706-188">必要に応じて、サイズ 16x16 px、20x20 px、24x24 px、40x40 px、48x48 px、および 64x64 px のアイコンを使用できます。</span><span class="sxs-lookup"><span data-stu-id="ab706-188">Optionally, you can also have icons of sizes 16x16 px, 20x20 px, 24x24 px, 40x40 px, 48x48 px, and 64x64 px.</span></span> <span data-ttu-id="ab706-189">Office、リボンのサイズとアプリケーション ウィンドウのサイズに基づいて使用するOffice決定します。</span><span class="sxs-lookup"><span data-stu-id="ab706-189">Office decides which icon to use based on the size of the ribbon and Office application window.</span></span> <span data-ttu-id="ab706-190">アイコン配列に次のオブジェクトを追加します。</span><span class="sxs-lookup"><span data-stu-id="ab706-190">Add the following objects to the icon array.</span></span> <span data-ttu-id="ab706-191">(ウィンドウとリボンのサイズが、グループのコントロールの少なくとも 1 つが表示されるのに十分な大きさの場合、グループ アイコンは表示されません。 </span><span class="sxs-lookup"><span data-stu-id="ab706-191">(If the window and ribbon sizes are large enough for at least one of the *controls* on the group to appear, then no group icon at all appears.</span></span> <span data-ttu-id="ab706-192">たとえば、Word **ウィンドウを縮小** して展開する場合は、Word リボンの [スタイル] グループを確認します)。このマークアップについて、次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="ab706-192">For an example, watch the **Styles** group on the Word ribbon as you shrink and expand the Word window.) About this markup, note:</span></span>

    - <span data-ttu-id="ab706-193">両方のプロパティが必要です。</span><span class="sxs-lookup"><span data-stu-id="ab706-193">Both the properties are required.</span></span>
    - <span data-ttu-id="ab706-194">プロパティ `size` の単位はピクセルです。</span><span class="sxs-lookup"><span data-stu-id="ab706-194">The `size` property unit of measure is pixels.</span></span> <span data-ttu-id="ab706-195">アイコンは常に正方形なので、数値は高さと幅の両方です。</span><span class="sxs-lookup"><span data-stu-id="ab706-195">Icons are always square, so the number is both the height and the width.</span></span>
    - <span data-ttu-id="ab706-196">プロパティ `sourceLocation` は、アイコンの完全な URL を指定します。</span><span class="sxs-lookup"><span data-stu-id="ab706-196">The `sourceLocation` property specifies the full URL to the icon.</span></span>

    > [!IMPORTANT]
    > <span data-ttu-id="ab706-197">開発から実稼働に移行する場合 (ドメインを localhost から contoso.com に変更するなど) ときに、アドインのマニフェスト内の URL を通常変更する必要があるのと同様に、コンテキスト タブ JSON の URL も変更する必要があります。</span><span class="sxs-lookup"><span data-stu-id="ab706-197">Just as you typically must change the URLs in the add-in's manifest when you move from development to production (such as changing the domain from localhost to contoso.com), you must also change the URLs in your contextual tabs JSON.</span></span>

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

1. <span data-ttu-id="ab706-198">この単純な進行中の例では、グループにはボタンが 1 つのみです。</span><span class="sxs-lookup"><span data-stu-id="ab706-198">In our simple ongoing example, the group has only a single button.</span></span> <span data-ttu-id="ab706-199">次のオブジェクトを配列の唯一のメンバーとして追加 `controls` します。</span><span class="sxs-lookup"><span data-stu-id="ab706-199">Add the following object as the only member of the `controls` array.</span></span> <span data-ttu-id="ab706-200">このマークアップについて、次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="ab706-200">About this markup, note:</span></span>

    - <span data-ttu-id="ab706-201">を除くすべての `enabled` プロパティが必要です。</span><span class="sxs-lookup"><span data-stu-id="ab706-201">All the properties, except `enabled`, are required.</span></span>
    - <span data-ttu-id="ab706-202">`type` コントロールの種類を指定します。</span><span class="sxs-lookup"><span data-stu-id="ab706-202">`type` specifies the type of control.</span></span> <span data-ttu-id="ab706-203">値には、"Button"、"Menu"、または "MobileButton" を指定できます。</span><span class="sxs-lookup"><span data-stu-id="ab706-203">The values can be "Button", "Menu", or "MobileButton".</span></span>
    - <span data-ttu-id="ab706-204">`id` 125 文字まで指定できます。</span><span class="sxs-lookup"><span data-stu-id="ab706-204">`id` can be up to 125 characters.</span></span> 
    - <span data-ttu-id="ab706-205">`actionId` は、配列で定義されているアクションの ID である必要 `actions` があります。</span><span class="sxs-lookup"><span data-stu-id="ab706-205">`actionId` must be the ID of an action defined in the `actions` array.</span></span> <span data-ttu-id="ab706-206">(このセクションの手順 1 を参照してください)。</span><span class="sxs-lookup"><span data-stu-id="ab706-206">(See step 1 of this section.)</span></span>
    - <span data-ttu-id="ab706-207">`label` は、ボタンのラベルとして機能するユーザーフレンドリーな文字列です。</span><span class="sxs-lookup"><span data-stu-id="ab706-207">`label` is a user-friendly string to serve as the label of the button.</span></span>
    - <span data-ttu-id="ab706-208">`superTip` は、豊富な形式のツール ヒントを表します。</span><span class="sxs-lookup"><span data-stu-id="ab706-208">`superTip` represents a rich form of tool tip.</span></span> <span data-ttu-id="ab706-209">プロパティと `title` プロパティ `description` の両方が必要です。</span><span class="sxs-lookup"><span data-stu-id="ab706-209">Both the `title` and `description` properties are required.</span></span>
    - <span data-ttu-id="ab706-210">`icon` ボタンのアイコンを指定します。</span><span class="sxs-lookup"><span data-stu-id="ab706-210">`icon` specifies the icons for the button.</span></span> <span data-ttu-id="ab706-211">グループ アイコンに関する前の説明もここでも適用されます。</span><span class="sxs-lookup"><span data-stu-id="ab706-211">The previous remarks about the group icon apply here too.</span></span>
    - <span data-ttu-id="ab706-212">`enabled` (省略可能) は、コンテキスト タブが表示されたら、ボタンを有効にするかどうかを指定します。</span><span class="sxs-lookup"><span data-stu-id="ab706-212">`enabled` (optional) specifies whether the button is enabled when the contextual tab appears starts up.</span></span> <span data-ttu-id="ab706-213">存在しない場合の既定値は `true` です。</span><span class="sxs-lookup"><span data-stu-id="ab706-213">The default if not present is `true`.</span></span> 

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
 
<span data-ttu-id="ab706-214">JSON BLOB の完全な例を次に示します。</span><span class="sxs-lookup"><span data-stu-id="ab706-214">The following is the complete example of the JSON blob:</span></span>

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

## <a name="register-the-contextual-tab-with-office-with-requestcreatecontrols"></a><span data-ttu-id="ab706-215">requestCreateControls でコンテキスト タブを Officeに登録する</span><span class="sxs-lookup"><span data-stu-id="ab706-215">Register the contextual tab with Office with requestCreateControls</span></span>

<span data-ttu-id="ab706-216">コンテキスト タブは[、Office.ribbon.requestCreateControls メソッドOffice呼](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls_tabDefinition_)び出すことによって、ユーザーに登録されます。</span><span class="sxs-lookup"><span data-stu-id="ab706-216">The contextual tab is registered with Office by calling the [Office.ribbon.requestCreateControls](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls_tabDefinition_) method.</span></span> <span data-ttu-id="ab706-217">これは通常、メソッドに割り当てられている関数またはメソッドで `Office.initialize` 行 `Office.onReady` われます。</span><span class="sxs-lookup"><span data-stu-id="ab706-217">This is typically done in either the function that is assigned to `Office.initialize` or with the `Office.onReady` method.</span></span> <span data-ttu-id="ab706-218">これらのメソッドとアドインの初期化の詳細については、「Initialize [your Office アドイン」を参照してください](../develop/initialize-add-in.md)。</span><span class="sxs-lookup"><span data-stu-id="ab706-218">For more about these methods and initializing the add-in, see [Initialize your Office Add-in](../develop/initialize-add-in.md).</span></span> <span data-ttu-id="ab706-219">ただし、初期化後にいつでもメソッドを呼び出す場合があります。</span><span class="sxs-lookup"><span data-stu-id="ab706-219">You can, however, call the method anytime after initialization.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="ab706-220">メソッド `requestCreateControls` は、アドインの特定のセッションで 1 回だけ呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="ab706-220">The `requestCreateControls` method can be called only once in a given session of an add-in.</span></span> <span data-ttu-id="ab706-221">再び呼び出された場合は、エラーがスローされます。</span><span class="sxs-lookup"><span data-stu-id="ab706-221">An error is thrown if it is called again.</span></span>

<span data-ttu-id="ab706-222">次に例を示します。</span><span class="sxs-lookup"><span data-stu-id="ab706-222">The following is an example.</span></span> <span data-ttu-id="ab706-223">JSON 文字列を JavaScript 関数に渡す前に、メソッドを使用して `JSON.parse` JavaScript オブジェクトに変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="ab706-223">Note that the JSON string must be converted to a JavaScript object with the `JSON.parse` method before it can be passed to a JavaScript function.</span></span>

```javascript
Office.onReady(async () => {
    const contextualTabJSON = ` ... `; // Assign the JSON string such as the one at the end of the preceding section.
    const contextualTab = JSON.parse(contextualTabJSON);
    await Office.ribbon.requestCreateControls(contextualTab);
});
```

## <a name="specify-the-contexts-when-the-tab-will-be-visible-with-requestupdate"></a><span data-ttu-id="ab706-224">requestUpdate でタブが表示されるコンテキストを指定する</span><span class="sxs-lookup"><span data-stu-id="ab706-224">Specify the contexts when the tab will be visible with requestUpdate</span></span>

<span data-ttu-id="ab706-225">通常、ユーザーが開始したイベントがアドイン コンテキストを変更すると、カスタム コンテキスト タブが表示されます。</span><span class="sxs-lookup"><span data-stu-id="ab706-225">Typically, a custom contextual tab should appear when a user-initiated event changes the add-in context.</span></span> <span data-ttu-id="ab706-226">グラフ (ブックの既定のワークシート) がアクティブ化されている場合にのみ、タブを表示する必要があるシナリオExcel考えます。</span><span class="sxs-lookup"><span data-stu-id="ab706-226">Consider a scenario in which the tab should be visible when, and only when, a chart (on the default worksheet of an Excel workbook) is activated.</span></span>

<span data-ttu-id="ab706-227">まず、ハンドラーを割り当てる。</span><span class="sxs-lookup"><span data-stu-id="ab706-227">Begin by assigning handlers.</span></span> <span data-ttu-id="ab706-228">これは、一般的に、ハンドラー (後の手順で作成) をワークシート内のすべてのグラフのイベントに割り当てる次の例のようにメソッド `Office.onReady` `onActivated` `onDeactivated` で行われます。</span><span class="sxs-lookup"><span data-stu-id="ab706-228">This is commonly done in the `Office.onReady` method as in the following example which assigns handlers (created in a later step) to the `onActivated` and `onDeactivated` events of all the charts in the worksheet.</span></span>

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

<span data-ttu-id="ab706-229">次に、ハンドラーを定義します。</span><span class="sxs-lookup"><span data-stu-id="ab706-229">Next, define the handlers.</span></span> <span data-ttu-id="ab706-230">次に示すのは、単純な例ですが、関数のより堅牢なバージョンについては、この記事の後半の `showDataTab` [「HostRestartNeeded](#handle-the-hostrestartneeded-error) エラーの処理」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="ab706-230">The following is a simple example of a `showDataTab`, but see [Handling the HostRestartNeeded error](#handle-the-hostrestartneeded-error) later in this article for a more robust version of the function.</span></span> <span data-ttu-id="ab706-231">このコードについては、以下の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="ab706-231">About this code, note:</span></span>

- <span data-ttu-id="ab706-232">Office では、リボンの状態を更新するタイミングが制御されます。</span><span class="sxs-lookup"><span data-stu-id="ab706-232">Office controls when it updates the state of the ribbon.</span></span> <span data-ttu-id="ab706-233">[Office.ribbon.requestUpdate](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestupdate-input-)メソッドは、更新要求をキューに入れられます。</span><span class="sxs-lookup"><span data-stu-id="ab706-233">The  [Office.ribbon.requestUpdate](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestupdate-input-) method queues a request to update.</span></span> <span data-ttu-id="ab706-234">リボンが実際に更新される場合ではなく、要求をキューに入れ次第、メソッド `Promise` はオブジェクトを解決します。</span><span class="sxs-lookup"><span data-stu-id="ab706-234">The method will resolve the `Promise` object as soon as it has queued the request, not when the ribbon actually updates.</span></span>
- <span data-ttu-id="ab706-235">メソッドのパラメーターは `requestUpdate` [RibbonUpdaterData](/javascript/api/office/office.ribbonupdaterdata) オブジェクトで、(1) *は JSON* で指定されたとおりにタブを ID で指定し、(2) はタブの表示を指定します。</span><span class="sxs-lookup"><span data-stu-id="ab706-235">The parameter for the `requestUpdate` method is a [RibbonUpdaterData](/javascript/api/office/office.ribbonupdaterdata) object that (1) specifies the tab by its ID *exactly as specified in the JSON* and (2) specifies visibility of the tab.</span></span>
- <span data-ttu-id="ab706-236">同じコンテキストで表示するカスタム コンテキスト タブが複数ある場合は、配列に追加のタブ オブジェクトを追加 `tabs` します。</span><span class="sxs-lookup"><span data-stu-id="ab706-236">If you have more than one custom contextual tab that should be visible in the same context, you simply add additional tab objects to the `tabs` array.</span></span>

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

<span data-ttu-id="ab706-237">タブを非表示にするハンドラーは、プロパティをに設定する以外は、ほぼ `visible` 同じです `false` 。</span><span class="sxs-lookup"><span data-stu-id="ab706-237">The handler to hide the tab is nearly identical, except that it sets the `visible` property back to `false`.</span></span>

<span data-ttu-id="ab706-238">JavaScript Officeには、オブジェクトを簡単に構築するためのいくつかのインターフェイス (型) も用意 `RibbonUpdateData` されています。</span><span class="sxs-lookup"><span data-stu-id="ab706-238">The Office JavaScript library also provides several interfaces (types) to make it easier to construct the`RibbonUpdateData` object.</span></span> <span data-ttu-id="ab706-239">TypeScript の `showDataTab` 関数を次に示します。これらの型を使用します。</span><span class="sxs-lookup"><span data-stu-id="ab706-239">The following is the `showDataTab` function in TypeScript and it makes use of these types.</span></span>

```typescript
const showDataTab = async () => {
    const myContextualTab: Office.Tab = {id: "CtxTab1", visible: true};
    const ribbonUpdater: Office.RibbonUpdaterData = { tabs: [ myContextualTab ]};
    await Office.ribbon.requestUpdate(ribbonUpdater);
}
```

### <a name="toggle-tab-visibility-and-the-enabled-status-of-a-button-at-the-same-time"></a><span data-ttu-id="ab706-240">タブの表示とボタンの有効な状態を同時に切り替える</span><span class="sxs-lookup"><span data-stu-id="ab706-240">Toggle tab visibility and the enabled status of a button at the same time</span></span>

<span data-ttu-id="ab706-241">このメソッドは、カスタム コンテキスト タブまたはカスタム コア タブのカスタム ボタンの有効または無効の状態を切り替える `requestUpdate` 場合にも使用されます。この詳細については、「Enable [and Disable Add-in Commands」を参照してください](disable-add-in-commands.md)。</span><span class="sxs-lookup"><span data-stu-id="ab706-241">The `requestUpdate` method is also used to toggle the enabled or disabled status of a custom button on either a custom contextual tab or a custom core tab. For details about this, see [Enable and Disable Add-in Commands](disable-add-in-commands.md).</span></span> <span data-ttu-id="ab706-242">タブの表示とボタンの有効な状態の両方を同時に変更するシナリオが考えられます。</span><span class="sxs-lookup"><span data-stu-id="ab706-242">There may be scenarios in which you want to change both the visibility of a tab and the enabled status of a button at the same time.</span></span> <span data-ttu-id="ab706-243">これは、1 回の呼び出しで行います `requestUpdate` 。</span><span class="sxs-lookup"><span data-stu-id="ab706-243">You can do this with a single call of `requestUpdate`.</span></span> <span data-ttu-id="ab706-244">次に、コンテキスト タブを表示すると同時に、コア タブのボタンが有効になっている例を示します。</span><span class="sxs-lookup"><span data-stu-id="ab706-244">The following is an example in which a button on a core tab is enabled at the same time as a contextual tab is made visible.</span></span>

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

<span data-ttu-id="ab706-245">次の例では、有効になっているボタンは、表示されるコンテキスト タブと全く同じです。</span><span class="sxs-lookup"><span data-stu-id="ab706-245">In the following example, the button that is enabled is on the very same contextual tab that is being made visible.</span></span>

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

## <a name="localizing-the-json-blob"></a><span data-ttu-id="ab706-246">JSON BLOB のローカライズ</span><span class="sxs-lookup"><span data-stu-id="ab706-246">Localizing the JSON blob</span></span>

<span data-ttu-id="ab706-247">渡される JSON BLOB は、カスタム コア タブのマニフェスト マークアップがローカライズされるのと同じ方法でローカライズされません (マニフェストからのローカライズの制御で `requestCreateControls` [説明します](../develop/localization.md#control-localization-from-the-manifest))。</span><span class="sxs-lookup"><span data-stu-id="ab706-247">The JSON blob that is passed to `requestCreateControls` is not localized the same way that the manifest markup for custom core tabs is localized (which is described at [Control localization from the manifest](../develop/localization.md#control-localization-from-the-manifest)).</span></span> <span data-ttu-id="ab706-248">代わりに、ローカライズは、ロケールごとに個別の JSON BLOB を使用して実行時に行う必要があります。</span><span class="sxs-lookup"><span data-stu-id="ab706-248">Instead, the localization must occur at runtime using distinct JSON blobs for each locale.</span></span> <span data-ttu-id="ab706-249">`switch` [Office.context.displayLanguage](/javascript/api/office/office.context#displayLanguage)プロパティをテストするステートメントを使用してください。</span><span class="sxs-lookup"><span data-stu-id="ab706-249">We suggest that you use a `switch` statement that tests the [Office.context.displayLanguage](/javascript/api/office/office.context#displayLanguage) property.</span></span> <span data-ttu-id="ab706-250">例を次に示します。</span><span class="sxs-lookup"><span data-stu-id="ab706-250">The following is an example:</span></span>

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

<span data-ttu-id="ab706-251">次に、次の例のように、コードで関数を呼び出して、渡されるローカライズされた BLOB `requestCreateControls` を取得します。</span><span class="sxs-lookup"><span data-stu-id="ab706-251">Then your code calls the function to get the localized blob that is passed to `requestCreateControls`, as in the following example:</span></span>

```javascript
var contextualTabJSON = GetContextualTabsJsonSupportedLocale();
```

## <a name="best-practices-for-custom-contextual-tabs"></a><span data-ttu-id="ab706-252">カスタム コンテキスト タブのベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="ab706-252">Best practices for custom contextual tabs</span></span>

### <a name="implement-an-alternate-ui-experience-when-custom-contextual-tabs-are-not-supported"></a><span data-ttu-id="ab706-253">カスタム コンテキスト タブがサポートされていない場合に、別の UI エクスペリエンスを実装する</span><span class="sxs-lookup"><span data-stu-id="ab706-253">Implement an alternate UI experience when custom contextual tabs are not supported</span></span>

<span data-ttu-id="ab706-254">プラットフォーム、アプリケーション、Office、Officeの組み合わせはサポートされていません `requestCreateControls` 。</span><span class="sxs-lookup"><span data-stu-id="ab706-254">Some combinations of platform, Office application, and Office build don't support `requestCreateControls`.</span></span> <span data-ttu-id="ab706-255">アドインは、これらの組み合わせの 1 つでアドインを実行しているユーザーに代替エクスペリエンスを提供するように設計されている必要があります。</span><span class="sxs-lookup"><span data-stu-id="ab706-255">Your add-in should be designed to provide an alternate experience to users who are running the add-in on one of those combinations.</span></span> <span data-ttu-id="ab706-256">次のセクションでは、フォールバック エクスペリエンスを提供する 2 つの方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="ab706-256">The following sections describe two ways of providing a fallback experience.</span></span>

#### <a name="use-noncontextual-tabs-or-controls"></a><span data-ttu-id="ab706-257">コンテキスト以外のタブまたはコントロールを使用する</span><span class="sxs-lookup"><span data-stu-id="ab706-257">Use noncontextual tabs or controls</span></span>

<span data-ttu-id="ab706-258">カスタム コンテキスト タブをサポートしないアプリケーションまたはプラットフォームでアドインが実行されている場合、カスタム コンテキスト タブを実装するアドインでフォールバック エクスペリエンスを作成するように設計されたマニフェスト要素 [、OverriddenByRibbonApi](../reference/manifest/overriddenbyribbonapi.md)があります。</span><span class="sxs-lookup"><span data-stu-id="ab706-258">There is a manifest element, [OverriddenByRibbonApi](../reference/manifest/overriddenbyribbonapi.md), that is designed to create a fallback experience in an add-in that implements custom contextual tabs when the add-in is running on an application or platform that doesn't support custom contextual tabs.</span></span> 

<span data-ttu-id="ab706-259">この要素を使用する最も簡単な方法は、アドインのカスタム コンテキスト タブのリボンカスタマイズを複製する 1 つ以上のカスタム コア タブ (つまり、コンテキストに依存しないカスタム タブ) をマニフェストで定義する方法です。</span><span class="sxs-lookup"><span data-stu-id="ab706-259">The simplest strategy for using this element is that you define in the manifest one or more custom core tabs (that is, *noncontextual* custom tabs) that duplicate the ribbon customizations of the custom contextual tabs in your add-in.</span></span> <span data-ttu-id="ab706-260">ただし `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` 、CustomTab の最初の子要素として [追加します](../reference/manifest/customtab.md)。</span><span class="sxs-lookup"><span data-stu-id="ab706-260">But you add `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` as the first child element of the [CustomTab](../reference/manifest/customtab.md).</span></span> <span data-ttu-id="ab706-261">その効果は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="ab706-261">The effect of doing so is the following:</span></span>

- <span data-ttu-id="ab706-262">カスタム コンテキスト タブをサポートするアプリケーションとプラットフォームでアドインが実行されている場合、カスタム コア タブはリボンに表示されません。</span><span class="sxs-lookup"><span data-stu-id="ab706-262">If the add-in runs on an application and platform that support custom contextual tabs, then the custom core tab won't appear on the ribbon.</span></span> <span data-ttu-id="ab706-263">代わりに、アドインがメソッドを呼び出す際に、カスタム コンテキスト タブが作成 `requestCreateControls` されます。</span><span class="sxs-lookup"><span data-stu-id="ab706-263">Instead, the custom contextual tab will be created when the add-in calls the `requestCreateControls` method.</span></span>
- <span data-ttu-id="ab706-264">アドインがサポートしていないアプリケーションまたはプラットフォームで実行されている場合は、カスタム コア タブが `requestCreateControls` リボンに表示されます。</span><span class="sxs-lookup"><span data-stu-id="ab706-264">If the add-in runs on an application or platform that *doesn't* support `requestCreateControls`, then the custom core tab does appear on the ribbon.</span></span>

<span data-ttu-id="ab706-265">この簡単な戦略の例を次に示します。</span><span class="sxs-lookup"><span data-stu-id="ab706-265">The following is an example of this simple strategy.</span></span>

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

<span data-ttu-id="ab706-266">この単純な戦略では、カスタム コンテキスト タブと子グループとコントロールをミラー化するカスタム コア タブを使用しますが、より複雑な戦略を使用できます。</span><span class="sxs-lookup"><span data-stu-id="ab706-266">This simple strategy uses a custom core tab that mirrors a custom contextual tab with it's child groups and controls, but you can use a more complex strategy.</span></span> <span data-ttu-id="ab706-267">要素は、Group 要素と Control 要素 (ボタンの種類とメニューの種類の両方) およびメニュー要素に (最初の) 子要素として `<OverriddenByRibbonApi>` 追加[](../reference/manifest/control.md#button-control)[](../reference/manifest/group.md)[](../reference/manifest/control.md)[](../reference/manifest/control.md#menu-dropdown-button-controls) `<Item>` することもできます。</span><span class="sxs-lookup"><span data-stu-id="ab706-267">The `<OverriddenByRibbonApi>` element can also be added as (the first) child element to the [Group](../reference/manifest/group.md) and [Control](../reference/manifest/control.md) elements (both [button type](../reference/manifest/control.md#button-control) and [menu type](../reference/manifest/control.md#menu-dropdown-button-controls)), and menu `<Item>` elements.</span></span> <span data-ttu-id="ab706-268">この事実により、コンテキスト タブに表示されるグループとコントロールを、さまざまなカスタム コア タブのさまざまなグループ、ボタン、およびメニューに分散できます。</span><span class="sxs-lookup"><span data-stu-id="ab706-268">This fact enables you to distribute the groups and controls that would otherwise appear on the contextual tab among various groups, buttons, and menus in various custom core tabs.</span></span> <span data-ttu-id="ab706-269">次に例を示します。</span><span class="sxs-lookup"><span data-stu-id="ab706-269">The following is an example.</span></span> <span data-ttu-id="ab706-270">カスタム コンテキスト タブがサポートされていない場合にのみ、カスタム コア タブに "MyButton" が表示されます。</span><span class="sxs-lookup"><span data-stu-id="ab706-270">Note that "MyButton" will appear on the custom core tab only when custom contextual tabs are not supported.</span></span> <span data-ttu-id="ab706-271">ただし、カスタム コンテキスト タブがサポートされているかどうかに関係なく、親グループとカスタム コア タブが表示されます。</span><span class="sxs-lookup"><span data-stu-id="ab706-271">But the parent group and custom core tab will appear regardless of whether custom contextual tabs are supported.</span></span>

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

<span data-ttu-id="ab706-272">その他の例については [、「OverriddenByRibbonApi」を参照してください](../reference/manifest/overriddenbyribbonapi.md)。</span><span class="sxs-lookup"><span data-stu-id="ab706-272">For more examples, see [OverriddenByRibbonApi](../reference/manifest/overriddenbyribbonapi.md).</span></span>

<span data-ttu-id="ab706-273">親タブ、グループ、またはメニューにマークが付いている場合、そのタブは表示されません。カスタム コンテキスト タブがサポートされていない場合、すべての子マークアップは `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` 無視されます。</span><span class="sxs-lookup"><span data-stu-id="ab706-273">When a parent tab, group, or menu is marked with `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>`, then it isn't visible, and all of it's child markup is ignored, when custom contextual tabs aren't supported.</span></span> <span data-ttu-id="ab706-274">したがって、これらの子要素に要素がある場合や、その値 `<OverriddenByRibbonApi>` が何かは関係ありません。</span><span class="sxs-lookup"><span data-stu-id="ab706-274">So, it doesn't matter if any of those child elements have the `<OverriddenByRibbonApi>` element or what its value is.</span></span> <span data-ttu-id="ab706-275">この意味は、メニュー項目、コントロール、またはグループをすべてのコンテキストで表示する必要がある場合、メニュー項目、コントロール、またはグループがマークされていない必要があるだけでなく、その親メニュー、グループ、およびタブもこの方法でマークする必要があるという意味です。 `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` </span><span class="sxs-lookup"><span data-stu-id="ab706-275">The implication of this is that if a menu item, control, or group must be visible in all contexts, then not only should it not be marked with `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>`, but *its ancestor menu, group, and tab must also not be marked this way*.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="ab706-276">タブ、グループ *、またはメニュー* のすべての子要素にマークを付けない `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` 。</span><span class="sxs-lookup"><span data-stu-id="ab706-276">Don't mark *all* of the child elements of a tab, group, or menu with `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>`.</span></span> <span data-ttu-id="ab706-277">前の段落で指定した理由で親要素がマークされている場合、これは `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` 意味をなします。</span><span class="sxs-lookup"><span data-stu-id="ab706-277">This is pointless if the parent element is marked with `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` for reasons given in the preceding paragraph.</span></span> <span data-ttu-id="ab706-278">さらに、親タブを使用しない (またはに設定する) 場合は、カスタム コンテキスト タブがサポートされているかどうかに関係なく、親が表示されますが、サポートされている場合は空になります `<OverriddenByRibbonApi>` `false` 。</span><span class="sxs-lookup"><span data-stu-id="ab706-278">Moreover, if you leave out the `<OverriddenByRibbonApi>` on the parent (or set it to `false`), then the parent will appear regardless of whether custom contextual tabs are supported, but it will be empty when they are supported.</span></span> <span data-ttu-id="ab706-279">したがって、カスタム コンテキスト タブがサポートされているときにすべての子要素が表示されない場合は、親にマークを付け、親のみを付け、 を指定します `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` 。</span><span class="sxs-lookup"><span data-stu-id="ab706-279">So, if all the child elements shouldn't appear when custom contextual tabs are supported, mark the parent, and only the parent, with `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>`.</span></span>

#### <a name="use-apis-that-show-or-hide-a-task-pane-in-specified-contexts"></a><span data-ttu-id="ab706-280">指定したコンテキストで作業ウィンドウを表示または非表示にする API を使用する</span><span class="sxs-lookup"><span data-stu-id="ab706-280">Use APIs that show or hide a task pane in specified contexts</span></span>

<span data-ttu-id="ab706-281">代わりに、アドインは、カスタム コンテキスト タブ上のコントロールの機能を複製する UI コントロールを含む作業ウィンドウ `<OverriddenByRibbonApi>` を定義できます。[次に、Office.addin.showAsTaskpane](/javascript/api/office/office.addin?view=common-js&preserve-view=true#showAsTaskpane__)メソッドと[Office.addin.hide](/javascript/api/office/office.addin?view=common-js&preserve-view=true#hide__)メソッドを使用して、サポートされている場合にのみコンテキスト タブが表示される場合にのみ作業ウィンドウを表示します。</span><span class="sxs-lookup"><span data-stu-id="ab706-281">As an alternative to `<OverriddenByRibbonApi>`, your add-in can define a task pane with UI controls that duplicate the functionality of the controls on a custom contextual tab. Then use the [Office.addin.showAsTaskpane](/javascript/api/office/office.addin?view=common-js&preserve-view=true#showAsTaskpane__) and [Office.addin.hide](/javascript/api/office/office.addin?view=common-js&preserve-view=true#hide__) methods to show the task pane when, and only when, the contextual tab would have been shown if it was supported.</span></span> <span data-ttu-id="ab706-282">これらのメソッドの使い方の詳細については、「アドインの作業ウィンドウを表示または非表示にするOffice[を参照してください](../develop/show-hide-add-in.md)。</span><span class="sxs-lookup"><span data-stu-id="ab706-282">For details on how to use these methods, see [Show or hide the task pane of your Office Add-in](../develop/show-hide-add-in.md).</span></span>

### <a name="handle-the-hostrestartneeded-error"></a><span data-ttu-id="ab706-283">HostRestartNeeded エラーの処理</span><span class="sxs-lookup"><span data-stu-id="ab706-283">Handle the HostRestartNeeded error</span></span>

<span data-ttu-id="ab706-284">一部のシナリオでは、Office はリボンを更新できず、エラーを返します。</span><span class="sxs-lookup"><span data-stu-id="ab706-284">In some scenarios, Office is unable to update the ribbon and will return an error.</span></span> <span data-ttu-id="ab706-285">たとえば、アドインがアップグレードされ、アップグレードされたアドインに異なるカスタム アドイン コマンドのセットがある場合は、Office アプリケーションを閉じてから、もう一度開く必要があります。</span><span class="sxs-lookup"><span data-stu-id="ab706-285">For example, if the add-in is upgraded and the upgraded add-in has a different set of custom add-in commands, then the Office application must be closed and reopened.</span></span> <span data-ttu-id="ab706-286">それまでの間、`requestUpdate` メソッドは `HostRestartNeeded` エラーを返します。</span><span class="sxs-lookup"><span data-stu-id="ab706-286">Until it is, the `requestUpdate` method will return the error `HostRestartNeeded`.</span></span> <span data-ttu-id="ab706-287">コードでこのエラーを処理する必要があります。</span><span class="sxs-lookup"><span data-stu-id="ab706-287">Your code should handle this error.</span></span> <span data-ttu-id="ab706-288">次に、方法の例を示します。</span><span class="sxs-lookup"><span data-stu-id="ab706-288">The following is an example of how.</span></span> <span data-ttu-id="ab706-289">この場合、`reportError` メソッドがユーザーにエラーを表示します。</span><span class="sxs-lookup"><span data-stu-id="ab706-289">In this case, the `reportError` method displays the error to the user.</span></span>

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
