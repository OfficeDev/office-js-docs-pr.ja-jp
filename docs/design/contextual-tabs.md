---
title: Office アドインでカスタムコンテキストタブを作成する
description: カスタムコンテキストタブを Office アドインに追加する方法について説明します。
ms.date: 11/20/2020
localization_priority: Normal
ms.openlocfilehash: 49a773aca0651b88c972c24a4cde0aa1e300d5e7
ms.sourcegitcommit: 6619e07cdfa68f9fa985febd5f03caf7aee57d5e
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 11/30/2020
ms.locfileid: "49505557"
---
# <a name="create-custom-contextual-tabs-in-office-add-ins-preview"></a><span data-ttu-id="1dc71-103">Office アドインでカスタムコンテキストタブを作成する (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="1dc71-103">Create custom contextual tabs in Office Add-ins (preview)</span></span>

<span data-ttu-id="1dc71-104">コンテキストタブは、office ドキュメントで指定されたイベントが発生したときにタブ行に表示される Office リボンの非表示タブコントロールです。</span><span class="sxs-lookup"><span data-stu-id="1dc71-104">A contextual tab is a hidden tab control in the Office ribbon that is displayed in the tab row when a specified event occurs in the Office document.</span></span> <span data-ttu-id="1dc71-105">たとえば、テーブルが選択されているときに、Excel のリボンに表示される [ **テーブルデザイン** ] タブ。</span><span class="sxs-lookup"><span data-stu-id="1dc71-105">For example, the **Table Design** tab that appears on the Excel ribbon when a table is selected.</span></span> <span data-ttu-id="1dc71-106">Office アドインにカスタムコンテキストタブを含めることができ、表示または非表示のタイミングを指定するには、表示を変更するイベントハンドラーを作成します。</span><span class="sxs-lookup"><span data-stu-id="1dc71-106">You can include custom contextual tabs in your Office add-in and specify when they are visible or hidden, by creating event handlers that change the visibility.</span></span> <span data-ttu-id="1dc71-107">(ただし、カスタムコンテキストタブはフォーカスの変更には応答しません)。</span><span class="sxs-lookup"><span data-stu-id="1dc71-107">(However, custom contextual tabs do not respond to focus changes.)</span></span>

> [!NOTE]
> <span data-ttu-id="1dc71-108">この記事は、以下のドキュメントについて既に理解していることを前提としています。</span><span class="sxs-lookup"><span data-stu-id="1dc71-108">This article assumes that you are familiar with the following documentation.</span></span> <span data-ttu-id="1dc71-109">最近、アドイン コマンド (カスタム メニュー項目とリボン ボタン) を使用してない場合は、ドキュメントをご確認ください。</span><span class="sxs-lookup"><span data-stu-id="1dc71-109">Please review it if you haven't worked with Add-in Commands (custom menu items and ribbon buttons) recently.</span></span>
>
> - [<span data-ttu-id="1dc71-110">アドイン コマンドの基本概念</span><span class="sxs-lookup"><span data-stu-id="1dc71-110">Basic concepts for Add-in Commands</span></span>](add-in-commands.md)

> [!IMPORTANT]
> <span data-ttu-id="1dc71-111">ユーザー設定のコンテキストタブはプレビューで表示されます。</span><span class="sxs-lookup"><span data-stu-id="1dc71-111">Custom contextual tabs are in preview.</span></span> <span data-ttu-id="1dc71-112">開発環境またはテスト環境で試してみることはできますが、運用アドインには追加しません。</span><span class="sxs-lookup"><span data-stu-id="1dc71-112">Please experiment with them in a development or testing environment but don't add them to a production add-in.</span></span>
>
> <span data-ttu-id="1dc71-113">現時点では、カスタムコンテキストタブは Excel でのみサポートされており、これらのプラットフォームでのみサポートされています。</span><span class="sxs-lookup"><span data-stu-id="1dc71-113">Custom contextual tabs are currently only supported on Excel and only on these platforms and builds:</span></span>
>
> - <span data-ttu-id="1dc71-114">Excel on Windows (Microsoft 365 のみ、永続的なライセンスではない): バージョン 2011 (ビルド 13426.20274)。</span><span class="sxs-lookup"><span data-stu-id="1dc71-114">Excel on Windows (Microsoft 365 only, not perpetual license): Version 2011 (Build 13426.20274).</span></span> <span data-ttu-id="1dc71-115">Microsoft 365 サブスクリプションは、以前は "月次 Channel (対象指定)" または "Insider 低速" と呼ばれていた [現在のチャネル (プレビュー)](https://insider.office.com/join/windows) にある必要があります。</span><span class="sxs-lookup"><span data-stu-id="1dc71-115">Your Microsoft 365 subscription may need to be on the [Current Channel (Preview)](https://insider.office.com/join/windows) formerly called "Monthly Channel (Targeted)" or "Insider Slow".</span></span>

> [!NOTE]
> <span data-ttu-id="1dc71-116">カスタムコンテキストタブは、次の要件セットをサポートするプラットフォームでのみ機能します。</span><span class="sxs-lookup"><span data-stu-id="1dc71-116">Custom contextual tabs work only on platforms that support the following requirement sets.</span></span> <span data-ttu-id="1dc71-117">要件セットの詳細とその使用方法については、「 [Office アプリケーションと API 要件を指定](../develop/specify-office-hosts-and-api-requirements.md)する」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="1dc71-117">For more about requirement sets and how to work with them, see [Specify Office applications and API requirements](../develop/specify-office-hosts-and-api-requirements.md).</span></span>
>
> - [<span data-ttu-id="1dc71-118">SharedRuntime 1.1</span><span class="sxs-lookup"><span data-stu-id="1dc71-118">SharedRuntime 1.1</span></span>](../reference/requirement-sets/shared-runtime-requirement-sets.md)

## <a name="behavior-of-custom-contextual-tabs"></a><span data-ttu-id="1dc71-119">カスタムコンテキストタブの動作</span><span class="sxs-lookup"><span data-stu-id="1dc71-119">Behavior of custom contextual tabs</span></span>

<span data-ttu-id="1dc71-120">カスタムコンテキストタブのユーザー操作は、組み込みの Office コンテキストタブのパターンに従います。</span><span class="sxs-lookup"><span data-stu-id="1dc71-120">The user experience for custom contextual tabs follows the pattern of built-in Office contextual tabs.</span></span> <span data-ttu-id="1dc71-121">配置カスタムコンテキストタブの基本的な原則を次に示します。</span><span class="sxs-lookup"><span data-stu-id="1dc71-121">The following are the basic principles for the placement custom contextual tabs:</span></span>

- <span data-ttu-id="1dc71-122">ユーザー設定のコンテキストタブが表示されている場合は、リボンの右端に表示されます。</span><span class="sxs-lookup"><span data-stu-id="1dc71-122">When a custom contextual tab is visible, it appears on the right end of the ribbon.</span></span>
- <span data-ttu-id="1dc71-123">1つ以上の組み込みコンテキストタブと、アドインのカスタムコンテキストタブが同時に表示されている場合は、カスタムコンテキストタブは常に、組み込みのコンテキストタブすべての右側にあります。</span><span class="sxs-lookup"><span data-stu-id="1dc71-123">If one or more built-in contextual tabs and one or more custom contextual tabs from add-ins are visible at the same time, the custom contextual tabs are always to the right of all of the built-in contextual tabs.</span></span>
- <span data-ttu-id="1dc71-124">アドインに複数のコンテキストタブがあり、複数のコンテキストタブが表示されている場合は、アドインで定義された順序で表示されます。</span><span class="sxs-lookup"><span data-stu-id="1dc71-124">If your add-in has more than one contextual tab and there are contexts in which more than one is visible, they appear in the order in which they are defined in your add-in.</span></span> <span data-ttu-id="1dc71-125">(方向は、Office の言語と同じ方向です。つまり、左から右に記述する言語では左から右ですが、右から左の言語では右から左)。定義方法の詳細については、 [タブに表示されるグループとコントロールを定義](#define-the-groups-and-controls-that-appear-on-the-tab) するを参照してください。</span><span class="sxs-lookup"><span data-stu-id="1dc71-125">(The direction is the same direction as the Office language; that is, is left-to-right in left-to-right languages, but right-to-left in right-to-left languages.) See [Define the groups and controls that appear on the tab](#define-the-groups-and-controls-that-appear-on-the-tab) for details about how you define them.</span></span>
- <span data-ttu-id="1dc71-126">複数のアドインにコンテキストタブがあり、特定のコンテキストに表示されている場合は、アドインが起動された順序で表示されます。</span><span class="sxs-lookup"><span data-stu-id="1dc71-126">If more than one add-in has a contextual tab that is visible in a specific context, then they appear in the order in which the add-ins were launched.</span></span>
- <span data-ttu-id="1dc71-127">カスタムの *コンテキスト* タブは、カスタムコアタブとは異なり、Office アプリケーションのリボンに永続的に追加されません。</span><span class="sxs-lookup"><span data-stu-id="1dc71-127">Custom *contextual* tabs, unlike custom core tabs, are not added permanently to the Office application's ribbon.</span></span> <span data-ttu-id="1dc71-128">これらは、アドインが実行されている Office ドキュメントにのみ存在します。</span><span class="sxs-lookup"><span data-stu-id="1dc71-128">They are present only in Office documents on which your add-in is running.</span></span>

## <a name="major-steps-for-including-a-contextual-tab-in-an-add-in"></a><span data-ttu-id="1dc71-129">アドインにコンテキストタブを含めるための主な手順</span><span class="sxs-lookup"><span data-stu-id="1dc71-129">Major steps for including a contextual tab in an add-in</span></span>

<span data-ttu-id="1dc71-130">アドインにカスタムコンテキストタブを含めるための主な手順を次に示します。</span><span class="sxs-lookup"><span data-stu-id="1dc71-130">The following are the major steps for including a custom contextual tab in an add-in:</span></span>

1. <span data-ttu-id="1dc71-131">共有ランタイムを使用するようにアドインを構成します。</span><span class="sxs-lookup"><span data-stu-id="1dc71-131">Configure the add-in to use a shared runtime.</span></span>
1. <span data-ttu-id="1dc71-132">タブと、その上に表示されるグループとコントロールを定義します。</span><span class="sxs-lookup"><span data-stu-id="1dc71-132">Define the tab and the groups and controls that appear on it.</span></span>
1. <span data-ttu-id="1dc71-133">コンテキストタブを Office に登録します。</span><span class="sxs-lookup"><span data-stu-id="1dc71-133">Register the contextual tab with Office.</span></span>
1. <span data-ttu-id="1dc71-134">タブが表示される状況を指定します。</span><span class="sxs-lookup"><span data-stu-id="1dc71-134">Specify the circumstances when the tab will be visible.</span></span>

## <a name="configure-the-add-in-to-use-a-shared-runtime"></a><span data-ttu-id="1dc71-135">共有ランタイムを使用するようにアドインを構成する</span><span class="sxs-lookup"><span data-stu-id="1dc71-135">Configure the add-in to use a shared runtime</span></span>

<span data-ttu-id="1dc71-136">カスタムコンテキストタブを追加するには、アドインで共有ランタイムを使用する必要があります。</span><span class="sxs-lookup"><span data-stu-id="1dc71-136">Adding custom contextual tabs requires your add-in to use the shared runtime.</span></span> <span data-ttu-id="1dc71-137">詳細については、「 [共有ランタイムを使用するようにアドインを構成する](../excel/configure-your-add-in-to-use-a-shared-runtime.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="1dc71-137">For more information, see [Configure an add-in to use a shared runtime](../excel/configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

## <a name="define-the-groups-and-controls-that-appear-on-the-tab"></a><span data-ttu-id="1dc71-138">タブに表示されるグループとコントロールを定義する</span><span class="sxs-lookup"><span data-stu-id="1dc71-138">Define the groups and controls that appear on the tab</span></span>

<span data-ttu-id="1dc71-139">マニフェストの XML で定義されているカスタムコアタブとは異なり、カスタムコンテキストタブは、JSON blob を使用して実行時に定義されます。</span><span class="sxs-lookup"><span data-stu-id="1dc71-139">Unlike custom core tabs, which are defined with XML in the manifest, custom contextual tabs are defined at runtime with a JSON blob.</span></span> <span data-ttu-id="1dc71-140">コードによって blob が JavaScript オブジェクトに解析され、そのオブジェクトが [requestCreateControls](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls-tabDefinition-) メソッドに渡されます。</span><span class="sxs-lookup"><span data-stu-id="1dc71-140">Your code parses the blob into a JavaScript object, and then passes the object to the [Office.ribbon.requestCreateControls](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls-tabDefinition-) method.</span></span> <span data-ttu-id="1dc71-141">カスタムコンテキストタブは、アドインが現在実行されているドキュメントにのみ存在します。</span><span class="sxs-lookup"><span data-stu-id="1dc71-141">Custom contextual tabs are only present in documents on which your add-in is currently running.</span></span> <span data-ttu-id="1dc71-142">これは、アドインがインストールされていて、別のドキュメントが開かれたときにそのまま表示される場合に、Office アプリケーションリボンに追加されるカスタムコアタブとは異なります。</span><span class="sxs-lookup"><span data-stu-id="1dc71-142">This is different from custom core tabs which are added to the Office application ribbon when the add-in is installed and remain present when another document is opened.</span></span> <span data-ttu-id="1dc71-143">また、この `requestCreateControls` メソッドは、アドインのセッションで一度だけ実行できます。</span><span class="sxs-lookup"><span data-stu-id="1dc71-143">Also, the `requestCreateControls` method can be run only once in a session of your add-in.</span></span> <span data-ttu-id="1dc71-144">再び呼び出された場合は、エラーがスローされます。</span><span class="sxs-lookup"><span data-stu-id="1dc71-144">If it is called again, an error is thrown.</span></span>

> [!NOTE]
> <span data-ttu-id="1dc71-145">JSON blob のプロパティとサブプロパティ (およびキー名) の構造は、マニフェスト XML 内の [Customtab](../reference/manifest/customtab.md) 要素とその子孫要素の構造とほぼ並行しています。</span><span class="sxs-lookup"><span data-stu-id="1dc71-145">The structure of the JSON blob's properties and subproperties (and the key names) is roughly parallel to the structure of the [CustomTab](../reference/manifest/customtab.md) element and its descendant elements in the manifest XML.</span></span>

<span data-ttu-id="1dc71-146">ここでは、コンテキストタブ JSON blob の詳細な手順を示します。</span><span class="sxs-lookup"><span data-stu-id="1dc71-146">We'll construct an example of a contextual tabs JSON blob step-by-step.</span></span> <span data-ttu-id="1dc71-147">(コンテキストタブ JSON の完全なスキーマは [dynamic-ribbon.schema.jsに](https://developer.microsoft.com/json-schemas/office-js/dynamic-ribbon.schema.json)あります。</span><span class="sxs-lookup"><span data-stu-id="1dc71-147">(The full schema for the contextual tab JSON is at [dynamic-ribbon.schema.json](https://developer.microsoft.com/json-schemas/office-js/dynamic-ribbon.schema.json).</span></span> <span data-ttu-id="1dc71-148">このリンクは、コンテキストタブの初期プレビュー期間では動作していない場合があります。</span><span class="sxs-lookup"><span data-stu-id="1dc71-148">This link may not be working in the early preview period for contextual tabs.</span></span> <span data-ttu-id="1dc71-149">リンクが機能していない場合は、 [ドラフト dynamic-ribbon.schema.js](https://github.com/OfficeDev/testing-assets/tree/master/jsonschema/dynamic-ribbon.schema.json)のスキーマの最新の下書きを見つけることができます。)Visual Studio Code で作業している場合は、このファイルを使用して IntelliSense を取得し、JSON を検証することができます。</span><span class="sxs-lookup"><span data-stu-id="1dc71-149">If the link is not working, you can find the latest draft of the schema at [draft dynamic-ribbon.schema.json](https://github.com/OfficeDev/testing-assets/tree/master/jsonschema/dynamic-ribbon.schema.json).) If you are working in Visual Studio Code, you can use this file to get IntelliSense and to validate your JSON.</span></span> <span data-ttu-id="1dc71-150">詳細については、「 [Visual Studio Code-json スキーマと設定を使用](https://code.visualstudio.com/docs/languages/json#_json-schemas-and-settings)して Json を編集する」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="1dc71-150">For more information, see [Editing JSON with Visual Studio Code - JSON schemas and settings](https://code.visualstudio.com/docs/languages/json#_json-schemas-and-settings).</span></span>


1. <span data-ttu-id="1dc71-151">最初に、という名前の2つの配列プロパティを使用して JSON 文字列を作成し `actions` `tabs` ます。</span><span class="sxs-lookup"><span data-stu-id="1dc71-151">Begin by creating a JSON string with two array properties named `actions` and `tabs`.</span></span> <span data-ttu-id="1dc71-152">配列は、 `actions` コンテキストタブのコントロールによって実行できるすべての関数の仕様です。配列は、 `tabs` 1 つ以上のコンテキストタブを定義します。 *最大値は 10* です。</span><span class="sxs-lookup"><span data-stu-id="1dc71-152">The `actions` array is a specification of all the functions that can be executed by controls on the contextual tab. The `tabs` array defines one or more contextual tabs, *up to a maximum of 10*.</span></span>

    ```json
    '{
      "actions": [

      ],
      "tabs": [

      ]
    }'
    ```

1. <span data-ttu-id="1dc71-153">このような状況に応じたタブの例では、1つのボタンだけが含まれるため、1つのアクションのみができます。</span><span class="sxs-lookup"><span data-stu-id="1dc71-153">This simple example of a contextual tab will have only a single button and, thus, only a single action.</span></span> <span data-ttu-id="1dc71-154">次のものを配列の唯一のメンバーとして追加し `actions` ます。</span><span class="sxs-lookup"><span data-stu-id="1dc71-154">Add the following as the only member of the `actions` array.</span></span> <span data-ttu-id="1dc71-155">このマークアップについて、次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="1dc71-155">About this markup, note:</span></span>

    - <span data-ttu-id="1dc71-156">`id`プロパティおよび `type` プロパティは必須です。</span><span class="sxs-lookup"><span data-stu-id="1dc71-156">The `id` and `type` properties are mandatory.</span></span>
    - <span data-ttu-id="1dc71-157">の値は、 `type` "ExecuteFunction" または "ShowTaskpane" のいずれかになります。</span><span class="sxs-lookup"><span data-stu-id="1dc71-157">The value of `type` can be either "ExecuteFunction" or "ShowTaskpane".</span></span>
    - <span data-ttu-id="1dc71-158">`functionName`プロパティは、の値がである場合にのみ使用され `type` `ExecuteFunction` ます。</span><span class="sxs-lookup"><span data-stu-id="1dc71-158">The `functionName` property is only used when the value of `type` is `ExecuteFunction`.</span></span> <span data-ttu-id="1dc71-159">これは、FunctionFile で定義されている関数の名前です。</span><span class="sxs-lookup"><span data-stu-id="1dc71-159">It is the name of a function defined in the FunctionFile.</span></span> <span data-ttu-id="1dc71-160">FunctionFile の詳細については、「 [アドインコマンドの基本的な概念](add-in-commands.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="1dc71-160">For more information about the FunctionFile, see [Basic concepts for Add-in Commands](add-in-commands.md).</span></span>
    - <span data-ttu-id="1dc71-161">この後の手順では、このアクションをコンテキストタブのボタンにマップします。</span><span class="sxs-lookup"><span data-stu-id="1dc71-161">In a later step, you will map this action to a button on the contextual tab.</span></span>

    ```json
    {
      "id": "executeWriteData",
      "type": "ExecuteFunction",
      "functionName": "writeData"
    }
   ```

1. <span data-ttu-id="1dc71-162">次のものを配列の唯一のメンバーとして追加し `tabs` ます。</span><span class="sxs-lookup"><span data-stu-id="1dc71-162">Add the following as the only member of the `tabs` array.</span></span> <span data-ttu-id="1dc71-163">このマークアップについて、次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="1dc71-163">About this markup, note:</span></span>

    - <span data-ttu-id="1dc71-164">`id` プロパティは必須です。</span><span class="sxs-lookup"><span data-stu-id="1dc71-164">The `id` property is required.</span></span> <span data-ttu-id="1dc71-165">簡単でわかりやすい ID を使用して、アドイン内のすべてのコンテキストタブにおいて一意にします。</span><span class="sxs-lookup"><span data-stu-id="1dc71-165">Use a brief, descriptive ID that is unique among all contextual tabs in your add-in.</span></span>
    - <span data-ttu-id="1dc71-166">`label` プロパティは必須です。</span><span class="sxs-lookup"><span data-stu-id="1dc71-166">The `label` property is required.</span></span> <span data-ttu-id="1dc71-167">これは、コンテキストタブのラベルとして機能する、ユーザーフレンドリな文字列です。</span><span class="sxs-lookup"><span data-stu-id="1dc71-167">It is a user-friendly string to serve as the label of the contextual tab.</span></span>
    - <span data-ttu-id="1dc71-168">`groups` プロパティは必須です。</span><span class="sxs-lookup"><span data-stu-id="1dc71-168">The `groups` property is required.</span></span> <span data-ttu-id="1dc71-169">タブに表示されるコントロールのグループを定義します。少なくとも1つのメンバーがあり *、20を超える* ことはできません。</span><span class="sxs-lookup"><span data-stu-id="1dc71-169">It defines the groups of controls that will appear on the tab. It must have at least one member *and no more than 20*.</span></span> <span data-ttu-id="1dc71-170">(カスタムコンテキストタブで使用できるコントロールの数にも制限があります。また、グループの数も制限されます。</span><span class="sxs-lookup"><span data-stu-id="1dc71-170">(There are also limits on the number of controls that you can have on a custom contextual tab and that will also constrain how many groups that you have.</span></span> <span data-ttu-id="1dc71-171">詳細については、次の手順を参照してください)。</span><span class="sxs-lookup"><span data-stu-id="1dc71-171">See the next step for more information.)</span></span>

    > [!NOTE]
    > <span data-ttu-id="1dc71-172">Tab オブジェクトには、 `visible` アドインの起動時にすぐにタブを表示するかどうかを指定するオプションのプロパティもあります。</span><span class="sxs-lookup"><span data-stu-id="1dc71-172">The tab object can also have an optional `visible` property that specifies whether the tab is visible immediately when the add-in starts up.</span></span> <span data-ttu-id="1dc71-173">通常、コンテキストタブは、ユーザーイベントによって表示がトリガーされる (ドキュメント内の一部の型のエンティティを選択するユーザーのような場合) ため、 `visible` プロパティは `false` 存在しない場合は既定値になります。</span><span class="sxs-lookup"><span data-stu-id="1dc71-173">Since contextual tabs are normally hidden until a user event triggers their visibility (such as the user selecting an entity of some type in the document), the `visible` property defaults to `false` when not present.</span></span> <span data-ttu-id="1dc71-174">後のセクションでは、イベントへの応答としてプロパティをに設定する方法を示し `true` ます。</span><span class="sxs-lookup"><span data-stu-id="1dc71-174">In a later section, we show how to set the property to `true` in response to an event.</span></span>

    ```json
    {
      "id": "CtxTab1",
      "label": "Data",
      "groups": [

      ]
    }
    ```

1. <span data-ttu-id="1dc71-175">この単純な例では、コンテキストタブには1つのグループしかありません。</span><span class="sxs-lookup"><span data-stu-id="1dc71-175">In the simple ongoing example, the contextual tab has only a single group.</span></span> <span data-ttu-id="1dc71-176">次のものを配列の唯一のメンバーとして追加し `groups` ます。</span><span class="sxs-lookup"><span data-stu-id="1dc71-176">Add the following as the only member of the `groups` array.</span></span> <span data-ttu-id="1dc71-177">このマークアップについて、次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="1dc71-177">About this markup, note:</span></span>

    - <span data-ttu-id="1dc71-178">すべてのプロパティが必要です。</span><span class="sxs-lookup"><span data-stu-id="1dc71-178">All the properties are required.</span></span>
    - <span data-ttu-id="1dc71-179">この `id` プロパティは、タブ内のすべてのグループ間で一意である必要があります。簡潔でわかりやすい ID を使用してください。</span><span class="sxs-lookup"><span data-stu-id="1dc71-179">The `id` property must be unique among all the groups in the tab. Use a brief, descriptive ID.</span></span>
    - <span data-ttu-id="1dc71-180">`label`は、グループのラベルとして機能する、ユーザーフレンドリな文字列です。</span><span class="sxs-lookup"><span data-stu-id="1dc71-180">The `label` is a user-friendly string to serve as the label of the group.</span></span>
    - <span data-ttu-id="1dc71-181">`icon`プロパティの値は、リボンと Office アプリケーションウィンドウのサイズに応じて、グループがリボン上に持つアイコンを指定するオブジェクトの配列です。</span><span class="sxs-lookup"><span data-stu-id="1dc71-181">The `icon` property's value is an array of objects that specify the icons that the group will have on the ribbon depending on the size of the ribbon and the Office application window.</span></span>
    - <span data-ttu-id="1dc71-182">`controls`プロパティの値は、グループ内のボタンやその他のコントロールを指定するオブジェクトの配列です。</span><span class="sxs-lookup"><span data-stu-id="1dc71-182">The `controls` property's value is an array of objects that specify the buttons and other controls in the group.</span></span> <span data-ttu-id="1dc71-183">*グループに* は、少なくとも1つの値が必要です。</span><span class="sxs-lookup"><span data-stu-id="1dc71-183">There must be at least one and *no more than 6 in a group*.</span></span>

    > [!IMPORTANT]
    > <span data-ttu-id="1dc71-184">*タブ全体のコントロールの合計数は20以下でなければなりません。*</span><span class="sxs-lookup"><span data-stu-id="1dc71-184">*The total number of controls on the whole tab can be no more than 20.*</span></span> <span data-ttu-id="1dc71-185">たとえば、6つのグループと2つのコントロールを持つ4つのグループを持つことができますが、それぞれ6つのコントロールを持つ4つのグループを持つことはできません。</span><span class="sxs-lookup"><span data-stu-id="1dc71-185">For example, you could have 3 groups with 6 controls each, and a fourth group with 2 controls, but you cannot have 4 groups with 6 controls each.</span></span>  

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

1. <span data-ttu-id="1dc71-186">すべてのグループには、少なくとも2つのサイズのアイコン、32x32 px、80x80 px が必要です。</span><span class="sxs-lookup"><span data-stu-id="1dc71-186">Every group must have an icon of at least two sizes, 32x32 px and 80x80 px.</span></span> <span data-ttu-id="1dc71-187">必要に応じて、サイズが16x16、20x20、24x24、40x40、48x48、64x64 のアイコンを設定することもできます。</span><span class="sxs-lookup"><span data-stu-id="1dc71-187">Optionally, you can also have icons of sizes 16x16, 20x20, 24x24, 40x40, 48x48 and 64x64.</span></span> <span data-ttu-id="1dc71-188">Office は、リボンおよび Office アプリケーションウィンドウのサイズに基づいて、どのアイコンを使用するかを決定します。</span><span class="sxs-lookup"><span data-stu-id="1dc71-188">Office decides which icon to use based on the size of the ribbon and Office application window.</span></span> <span data-ttu-id="1dc71-189">アイコン配列に次のオブジェクトを追加します。</span><span class="sxs-lookup"><span data-stu-id="1dc71-189">Add the following objects to the icon array.</span></span> <span data-ttu-id="1dc71-190">(ウィンドウ上の少なくとも1つの *コントロール* が表示されるようにウィンドウとリボンのサイズが大きい場合、[グループ] アイコンはまったく表示されません。</span><span class="sxs-lookup"><span data-stu-id="1dc71-190">(If the window and ribbon sizes are large enough for at least one of the *controls* on the group to appear, then no group icon at all appears.</span></span> <span data-ttu-id="1dc71-191">例については、word のウィンドウを縮小して展開するときに、Word のリボンの [ **スタイル** ] グループを見てください。このマークアップについて、次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="1dc71-191">For an example, watch the **Styles** group on the Word ribbon as you shrink and expand the Word window.) About this markup, note:</span></span>

    - <span data-ttu-id="1dc71-192">両方のプロパティが必要です。</span><span class="sxs-lookup"><span data-stu-id="1dc71-192">Both the properties are required.</span></span>
    - <span data-ttu-id="1dc71-193">`size`プロパティの測定単位はピクセルです。</span><span class="sxs-lookup"><span data-stu-id="1dc71-193">The `size` property unit of measure is pixels.</span></span> <span data-ttu-id="1dc71-194">アイコンは常に正方形なので、数字は高さと幅の両方です。</span><span class="sxs-lookup"><span data-stu-id="1dc71-194">Icons are always square, so the number is both the height and the width.</span></span>
    - <span data-ttu-id="1dc71-195">この `sourceLocation` プロパティは、アイコンへの完全な URL を指定します。</span><span class="sxs-lookup"><span data-stu-id="1dc71-195">The `sourceLocation` property specifies the full URL to the icon.</span></span>

    > [!IMPORTANT]
    > <span data-ttu-id="1dc71-196">開発環境から運用環境に移行するときに、通常、アドインのマニフェストの Url を変更する必要があるのと同様に、コンテキストタブ JSON の Url も変更する必要があります。</span><span class="sxs-lookup"><span data-stu-id="1dc71-196">Just as you typically must change the URLs in the add-in's manifest when you move from development to production (such as changing the domain from localhost to contoso.com), you must also change the URLs in your contextual tabs JSON.</span></span>

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

1. <span data-ttu-id="1dc71-197">この簡単な例では、グループにボタンが1つしかありません。</span><span class="sxs-lookup"><span data-stu-id="1dc71-197">In our simple ongoing example, the group has only a single button.</span></span> <span data-ttu-id="1dc71-198">次のオブジェクトを配列の唯一のメンバーとして追加し `controls` ます。</span><span class="sxs-lookup"><span data-stu-id="1dc71-198">Add the following object as the only member of the `controls` array.</span></span> <span data-ttu-id="1dc71-199">このマークアップについて、次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="1dc71-199">About this markup, note:</span></span>

    - <span data-ttu-id="1dc71-200">以外のすべてのプロパティ `enabled` は必須です。</span><span class="sxs-lookup"><span data-stu-id="1dc71-200">All the properties, except `enabled`, are required.</span></span>
    - <span data-ttu-id="1dc71-201">`type` コントロールの種類を指定します。</span><span class="sxs-lookup"><span data-stu-id="1dc71-201">`type` specifies the type of control.</span></span> <span data-ttu-id="1dc71-202">値には、"Button"、"Menu"、または "MobileButton" を指定できます。</span><span class="sxs-lookup"><span data-stu-id="1dc71-202">The values can be "Button", "Menu", or "MobileButton".</span></span>
    - <span data-ttu-id="1dc71-203">`id` 最大125文字を使用できます。</span><span class="sxs-lookup"><span data-stu-id="1dc71-203">`id` can be up to 125 characters.</span></span> 
    - <span data-ttu-id="1dc71-204">`actionId` は、配列で定義されているアクションの ID である必要があり `actions` ます。</span><span class="sxs-lookup"><span data-stu-id="1dc71-204">`actionId` must be the ID of an action defined in the `actions` array.</span></span> <span data-ttu-id="1dc71-205">(このセクションの手順1を参照してください)。</span><span class="sxs-lookup"><span data-stu-id="1dc71-205">(See step 1 of this section.)</span></span>
    - <span data-ttu-id="1dc71-206">`label` は、ボタンのラベルとして機能するわかりやすい文字列です。</span><span class="sxs-lookup"><span data-stu-id="1dc71-206">`label` is a user-friendly string to serve as the label of the button.</span></span>
    - <span data-ttu-id="1dc71-207">`superTip` ツールヒントの豊富な形式を表します。</span><span class="sxs-lookup"><span data-stu-id="1dc71-207">`superTip` represents a rich form of tool tip.</span></span> <span data-ttu-id="1dc71-208">`title`プロパティとプロパティの両方 `description` が必要です。</span><span class="sxs-lookup"><span data-stu-id="1dc71-208">Both the `title` and `description` properties are required.</span></span>
    - <span data-ttu-id="1dc71-209">`icon` ボタンのアイコンを指定します。</span><span class="sxs-lookup"><span data-stu-id="1dc71-209">`icon` specifies the icons for the button.</span></span> <span data-ttu-id="1dc71-210">グループアイコンに関する上記の解説も適用されます。</span><span class="sxs-lookup"><span data-stu-id="1dc71-210">The previous remarks about the group icon apply here too.</span></span>
    - <span data-ttu-id="1dc71-211">`enabled` (省略可能) [操作] タブが表示されたときにボタンを有効にするかどうかを指定します。</span><span class="sxs-lookup"><span data-stu-id="1dc71-211">`enabled` (optional) specifies whether the button is enabled when the contextual tab appears starts up.</span></span> <span data-ttu-id="1dc71-212">存在しない場合の既定値は、 `true` です。</span><span class="sxs-lookup"><span data-stu-id="1dc71-212">The default if not present is `true`.</span></span> 

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
 
<span data-ttu-id="1dc71-213">JSON blob の完全な例を次に示します。</span><span class="sxs-lookup"><span data-stu-id="1dc71-213">The following is the complete example of the JSON blob:</span></span>

```json
'{
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
}'
```

## <a name="register-the-contextual-tab-with-office-with-requestcreatecontrols"></a><span data-ttu-id="1dc71-214">RequestCreateControls を使用して Office にコンテキストタブを登録する</span><span class="sxs-lookup"><span data-stu-id="1dc71-214">Register the contextual tab with Office with requestCreateControls</span></span>

<span data-ttu-id="1dc71-215">コンテキストタブは、 [requestCreateControls](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls_tabDefinition_) メソッドを呼び出すことによって、office に登録されています。</span><span class="sxs-lookup"><span data-stu-id="1dc71-215">The contextual tab is registered with Office by calling the [Office.ribbon.requestCreateControls](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls_tabDefinition_) method.</span></span> <span data-ttu-id="1dc71-216">これは、通常、またはメソッドによって割り当てられた関数によって実行され `Office.initialize` `Office.onReady` ます。</span><span class="sxs-lookup"><span data-stu-id="1dc71-216">This is typically done in either the function that is assigned to `Office.initialize` or with the `Office.onReady` method.</span></span> <span data-ttu-id="1dc71-217">これらのメソッドとアドインの初期化の詳細については、「 [Office アドインを初期化する](../develop/initialize-add-in.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="1dc71-217">For more about these methods and initializing the add-in, see [Initialize your Office Add-in](../develop/initialize-add-in.md).</span></span> <span data-ttu-id="1dc71-218">ただし、初期化後はいつでもメソッドを呼び出すことができます。</span><span class="sxs-lookup"><span data-stu-id="1dc71-218">You can, however, call the method anytime after initialization.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="1dc71-219">この `requestCreateControls` メソッドは、アドインの特定のセッションで1回だけ呼び出すことができます。</span><span class="sxs-lookup"><span data-stu-id="1dc71-219">The `requestCreateControls` method can be called only once in a given session of an add-in.</span></span> <span data-ttu-id="1dc71-220">再び呼び出された場合、エラーがスローされます。</span><span class="sxs-lookup"><span data-stu-id="1dc71-220">An error is thrown if it is called again.</span></span>

<span data-ttu-id="1dc71-221">次に例を示します。</span><span class="sxs-lookup"><span data-stu-id="1dc71-221">The following is an example.</span></span> <span data-ttu-id="1dc71-222">JSON 文字列は、 `JSON.parse` javascript 関数に渡す前に、メソッドを使用して javascript オブジェクトに変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="1dc71-222">Note that the JSON string must be converted to a JavaScript object with the `JSON.parse` method before it can be passed to a JavaScript function.</span></span>

```javascript
Office.onReady(async () => {
    const contextualTabJSON = ' ... '; // Assign the JSON string such as the one at the end of the preceding section.
    const contextualTab = JSON.parse(contextualTabJSON);
    await Office.ribbon.requestCreateControls(contextualTab);
});
```

## <a name="specify-the-contexts-when-the-tab-will-be-visible-with-requestupdate"></a><span data-ttu-id="1dc71-223">RequestUpdate でタブが表示されるときのコンテキストを指定する</span><span class="sxs-lookup"><span data-stu-id="1dc71-223">Specify the contexts when the tab will be visible with requestUpdate</span></span>

<span data-ttu-id="1dc71-224">通常、ユーザーが開始したイベントによってアドインコンテキストが変更されたときに、ユーザー設定のコンテキストタブが表示されるようにする必要があります。</span><span class="sxs-lookup"><span data-stu-id="1dc71-224">Typically, a custom contextual tab should appear when a user-initiated event changes the add-in context.</span></span> <span data-ttu-id="1dc71-225">Excel ブックの既定のワークシートにあるグラフがアクティブ化されたときにのみタブが表示されるシナリオを考えてみます。</span><span class="sxs-lookup"><span data-stu-id="1dc71-225">Consider a scenario in which the tab should be visible when, and only when, a chart (on the default worksheet of an Excel workbook) is activated.</span></span>

<span data-ttu-id="1dc71-226">最初に、ハンドラーを割り当てます。</span><span class="sxs-lookup"><span data-stu-id="1dc71-226">Begin by assigning handlers.</span></span> <span data-ttu-id="1dc71-227">これは通常、次の例のようにメソッドで行われ `Office.onReady` ます。この例では、ハンドラー (後の手順で作成したもの) を、 `onActivated` `onDeactivated` ワークシート内のすべてのグラフのイベントに割り当てます。</span><span class="sxs-lookup"><span data-stu-id="1dc71-227">This is commonly done in the `Office.onReady` method as in the following example which assigns handlers (created in a later step) to the `onActivated` and `onDeactivated` events of all the charts in the worksheet.</span></span>

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

<span data-ttu-id="1dc71-228">次に、ハンドラーを定義します。</span><span class="sxs-lookup"><span data-stu-id="1dc71-228">Next, define the handlers.</span></span> <span data-ttu-id="1dc71-229">の簡単な例を次に示し `showDataTab` ますが、より堅牢なバージョンの関数については、この記事で後述する「 [エラー処理](#error-handling) 」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="1dc71-229">The following is a simple example of a `showDataTab`, but see [Error Handling](#error-handling) later in this article for a more robust version of the function.</span></span> <span data-ttu-id="1dc71-230">このコードについては、以下の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="1dc71-230">About this code, note:</span></span>

- <span data-ttu-id="1dc71-231">Office では、リボンの状態を更新するタイミングが制御されます。</span><span class="sxs-lookup"><span data-stu-id="1dc71-231">Office controls when it updates the state of the ribbon.</span></span> <span data-ttu-id="1dc71-232">[Office の更新](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestupdate-input-)要求をキューに入れます。</span><span class="sxs-lookup"><span data-stu-id="1dc71-232">The  [Office.ribbon.requestUpdate](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestupdate-input-) method queues a request to update.</span></span> <span data-ttu-id="1dc71-233">このメソッドは、 `Promise` リボンが実際に更新されるときではなく、要求がキューに入った直後にオブジェクトを解決します。</span><span class="sxs-lookup"><span data-stu-id="1dc71-233">The method will resolve the `Promise` object as soon as it has queued the request, not when the ribbon actually updates.</span></span>
- <span data-ttu-id="1dc71-234">メソッドのパラメーター `requestUpdate` は、 *JSON で指定され* た ID を使用してタブを指定する [RibbonUpdaterData](/javascript/api/office/office.ribbonupdaterdata)オブジェクトであり、(2) はタブの表示を指定します。</span><span class="sxs-lookup"><span data-stu-id="1dc71-234">The parameter for the `requestUpdate` method is a [RibbonUpdaterData](/javascript/api/office/office.ribbonupdaterdata) object that (1) specifies the tab by its ID *exactly as specified in the JSON* and (2) specifies visibility of the tab.</span></span>
- <span data-ttu-id="1dc71-235">同じコンテキストに表示する必要があるカスタムコンテキストタブが複数ある場合は、単に配列に追加の tab オブジェクトを追加し `tabs` ます。</span><span class="sxs-lookup"><span data-stu-id="1dc71-235">If you have more than one custom contextual tab that should be visible in the same context, you simply add additional tab objects to the `tabs` array.</span></span>

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

<span data-ttu-id="1dc71-236">タブを非表示にするハンドラーはほぼ同じですが、プロパティがに戻される点が異なり `visible` `false` ます。</span><span class="sxs-lookup"><span data-stu-id="1dc71-236">The handler to hide the tab is nearly identical, except that it sets the `visible` property back to `false`.</span></span>

<span data-ttu-id="1dc71-237">Office JavaScript ライブラリには、オブジェクトを簡単に作成できるように、いくつかのインターフェイス (型) も用意されて `RibbonUpdateData` います。</span><span class="sxs-lookup"><span data-stu-id="1dc71-237">The Office JavaScript library also provides several interfaces (types) to make it easier to construct the`RibbonUpdateData` object.</span></span> <span data-ttu-id="1dc71-238">次に示すのは `showDataTab` TypeScript の関数で、これらの型を使用します。</span><span class="sxs-lookup"><span data-stu-id="1dc71-238">The following is the `showDataTab` function in TypeScript and it makes use of these types.</span></span>

```typescript
const showDataTab = async () => {
    const myContextualTab: Tab = {id: "CtxTab1", visible: true};
    const ribbonUpdater: RibbonUpdaterData = { tabs: [ myContextualTab ]};
    await Office.ribbon.requestUpdate(ribbonUpdater);
}
```

### <a name="toggle-tab-visibility-and-the-enabled-status-of-a-button-at-the-same-time"></a><span data-ttu-id="1dc71-239">ボタンの表示/非表示の状態を同時に切り替えます。</span><span class="sxs-lookup"><span data-stu-id="1dc71-239">Toggle tab visibility and the enabled status of a button at the same time</span></span>

<span data-ttu-id="1dc71-240">この `requestUpdate` メソッドは、カスタムコンテキストタブまたはカスタムコアタブのいずれかで、カスタムボタンの有効または無効の状態を切り替えるためにも使用されます。詳細については、「 [アドインコマンドを有効または無効](disable-add-in-commands.md)にする」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="1dc71-240">The `requestUpdate` method is also used to toggle the enabled or disabled status of a custom button on either a custom contextual tab or a custom core tab. For details about this, see [Enable and Disable Add-in Commands](disable-add-in-commands.md).</span></span> <span data-ttu-id="1dc71-241">タブの表示とボタンの有効な状態の両方を同時に変更するシナリオがある場合があります。</span><span class="sxs-lookup"><span data-stu-id="1dc71-241">There may be scenarios in which you want to change both the visibility of a tab and the enabled status of a button at the same time.</span></span> <span data-ttu-id="1dc71-242">これは、の1回の呼び出しで行うことができ `requestUpdate` ます。</span><span class="sxs-lookup"><span data-stu-id="1dc71-242">You can do this with a single call of `requestUpdate`.</span></span> <span data-ttu-id="1dc71-243">次の例は、コンテキストタブが表示されているときに、コアタブのボタンが有効になっています。</span><span class="sxs-lookup"><span data-stu-id="1dc71-243">The following is an example in which a button on a core tab is enabled at the same time as a contextual tab is made visible.</span></span>

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

<span data-ttu-id="1dc71-244">次の例では、有効になっているボタンは、表示されているのと同じコンテキストタブにあります。</span><span class="sxs-lookup"><span data-stu-id="1dc71-244">In the following example, the button that is enabled is on the very same contextual tab that is being made visible.</span></span>

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

## <a name="error-handling"></a><span data-ttu-id="1dc71-245">エラー処理</span><span class="sxs-lookup"><span data-stu-id="1dc71-245">Error handling</span></span>

<span data-ttu-id="1dc71-246">一部のシナリオでは、Office はリボンを更新できず、エラーを返します。</span><span class="sxs-lookup"><span data-stu-id="1dc71-246">In some scenarios, Office is unable to update the ribbon and will return an error.</span></span> <span data-ttu-id="1dc71-247">たとえば、アドインがアップグレードされ、アップグレードされたアドインに異なるカスタム アドイン コマンドのセットがある場合は、Office アプリケーションを閉じてから、もう一度開く必要があります。</span><span class="sxs-lookup"><span data-stu-id="1dc71-247">For example, if the add-in is upgraded and the upgraded add-in has a different set of custom add-in commands, then the Office application must be closed and reopened.</span></span> <span data-ttu-id="1dc71-248">それまでの間、`requestUpdate` メソッドは `HostRestartNeeded` エラーを返します。</span><span class="sxs-lookup"><span data-stu-id="1dc71-248">Until it is, the `requestUpdate` method will return the error `HostRestartNeeded`.</span></span> <span data-ttu-id="1dc71-249">このエラーの処理方法の例を次に示します。</span><span class="sxs-lookup"><span data-stu-id="1dc71-249">The following is an example of how to handle this error.</span></span> <span data-ttu-id="1dc71-250">この場合、`reportError` メソッドがユーザーにエラーを表示します。</span><span class="sxs-lookup"><span data-stu-id="1dc71-250">In this case, the `reportError` method displays the error to the user.</span></span>

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
