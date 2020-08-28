---
title: アドイン コマンドを有効または無効にする
description: Office Web アドインのカスタム リボン ボタンとメニュー項目の有効または無効の状態を変更する方法について説明します。
ms.date: 08/26/2020
localization_priority: Normal
ms.openlocfilehash: 54bfa06a3acfbea561d20a1b327f093429d725fc
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/28/2020
ms.locfileid: "47292975"
---
# <a name="enable-and-disable-add-in-commands"></a><span data-ttu-id="1d9eb-103">アドイン コマンドを有効または無効にする</span><span class="sxs-lookup"><span data-stu-id="1d9eb-103">Enable and Disable Add-in Commands</span></span>

<span data-ttu-id="1d9eb-104">アドインの一部の機能を特定のコンテキストでのみ使用可能にする必要がある場合、カスタム アドイン コマンドをプログラムで有効または無効にすることができます。</span><span class="sxs-lookup"><span data-stu-id="1d9eb-104">When some functionality in your add-in should only be available in certain contexts, you can programmatically enable or disable your custom Add-in Commands.</span></span> <span data-ttu-id="1d9eb-105">たとえば、表の見出しを変更する関数は、カーソルが表の中にある場合にのみ有効にする必要があります。</span><span class="sxs-lookup"><span data-stu-id="1d9eb-105">For example, a function that changes the header of a table should only be enabled when the cursor is in a table.</span></span>

<span data-ttu-id="1d9eb-106">また、Office クライアントアプリケーションを開いたときにコマンドを有効にするか無効にするかを指定することもできます。</span><span class="sxs-lookup"><span data-stu-id="1d9eb-106">You can also specify whether the command is enabled or disabled when the Office client application opens.</span></span>

> [!NOTE]
> <span data-ttu-id="1d9eb-107">この記事は、以下のドキュメントについて既に理解していることを前提としています。</span><span class="sxs-lookup"><span data-stu-id="1d9eb-107">This article assumes that you are familiar with the following documentation.</span></span> <span data-ttu-id="1d9eb-108">最近、アドイン コマンド (カスタム メニュー項目とリボン ボタン) を使用してない場合は、ドキュメントをご確認ください。</span><span class="sxs-lookup"><span data-stu-id="1d9eb-108">Please review it if you haven't worked with Add-in Commands (custom menu items and ribbon buttons) recently.</span></span>
>
> - [<span data-ttu-id="1d9eb-109">アドイン コマンドの基本概念</span><span class="sxs-lookup"><span data-stu-id="1d9eb-109">Basic concepts for Add-in Commands</span></span>](add-in-commands.md)

## <a name="office-application-and-platform-support-only"></a><span data-ttu-id="1d9eb-110">Office アプリケーションとプラットフォームのサポートのみ</span><span class="sxs-lookup"><span data-stu-id="1d9eb-110">Office application and platform support only</span></span>

<span data-ttu-id="1d9eb-111">この記事に記載されている Api は、Excel でのみ使用できます。また、Windows および Office on Mac 上の Office でのみ利用できます。</span><span class="sxs-lookup"><span data-stu-id="1d9eb-111">The APIs described in this article are only available in Excel and only on Office on Windows and Office on Mac.</span></span>

### <a name="test-for-platform-support-with-requirement-sets"></a><span data-ttu-id="1d9eb-112">要件セットを使用したプラットフォーム サポートのテスト</span><span class="sxs-lookup"><span data-stu-id="1d9eb-112">Test for platform support with requirement sets</span></span>

<span data-ttu-id="1d9eb-113">要件セットは、API メンバーの名前付きグループです。</span><span class="sxs-lookup"><span data-stu-id="1d9eb-113">Requirement sets are named groups of API members.</span></span> <span data-ttu-id="1d9eb-114">Office アドインは、マニフェストで指定されている要件セットを使用するか、ランタイムチェックを使用して、Office アプリケーションとプラットフォームの組み合わせがアドインに必要な Api をサポートしているかどうかを判断します。</span><span class="sxs-lookup"><span data-stu-id="1d9eb-114">Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office application and platform combination supports APIs that an add-in needs.</span></span> <span data-ttu-id="1d9eb-115">詳細については、「 [Office のバージョンと要件セット](../develop/office-versions-and-requirement-sets.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="1d9eb-115">For more information, see [Office versions and requirement sets](../develop/office-versions-and-requirement-sets.md).</span></span>

<span data-ttu-id="1d9eb-116">Enable/disable Api は、 [ribbonapi 1.1](../reference/requirement-sets/ribbon-api-requirement-sets.md) 要件セットに属しています。</span><span class="sxs-lookup"><span data-stu-id="1d9eb-116">The enable/disable APIs belong to the [RibbonApi 1.1](../reference/requirement-sets/ribbon-api-requirement-sets.md) requirement set.</span></span>

> [!NOTE]
> <span data-ttu-id="1d9eb-117">**Ribbonapi 1.1**要件セットはマニフェストでまだサポートされていないため、マニフェストのセクションで指定することはできません `<Requirements>` 。</span><span class="sxs-lookup"><span data-stu-id="1d9eb-117">The **RibbonApi 1.1** requirement set is not yet supported in the manifest, so you cannot specify it in the manifest's `<Requirements>` section.</span></span> <span data-ttu-id="1d9eb-118">サポートをテストするには、コードがを呼び出す必要があり `Office.context.requirements.isSetSupported('RibbonApi', '1.1')` ます。</span><span class="sxs-lookup"><span data-stu-id="1d9eb-118">To test for support, your code should call `Office.context.requirements.isSetSupported('RibbonApi', '1.1')`.</span></span> <span data-ttu-id="1d9eb-119">呼び出しが戻る *場合に限り*、コードで `true` Enable/disable api を呼び出すことができます。</span><span class="sxs-lookup"><span data-stu-id="1d9eb-119">If, *and only if*, that call returns `true`, your code can call the enable/disable APIs.</span></span> <span data-ttu-id="1d9eb-120">を呼び出した場合 `isSetSupported` は `false` 、すべてのカスタムアドインコマンドが常に有効になります。</span><span class="sxs-lookup"><span data-stu-id="1d9eb-120">If the call of `isSetSupported` returns `false`, then all custom add-in commands are enabled all of the time.</span></span> <span data-ttu-id="1d9eb-121">**Ribbonapi 1.1**要件セットがサポートされていない場合にどのように動作するかを考慮するには、運用アドインとアプリ内の手順を設計する必要があります。</span><span class="sxs-lookup"><span data-stu-id="1d9eb-121">You must design your production add-in, and any in-app instructions, to take account of how it will work when the **RibbonApi 1.1** requirement set is not supported.</span></span> <span data-ttu-id="1d9eb-122">の使用法の詳細と例については `isSetSupported` 、「 [Office アプリケーションと API 要件を指定](../develop/specify-office-hosts-and-api-requirements.md)する」を参照してください。特に、 [JavaScript コードでランタイムチェックを使用](../develop/specify-office-hosts-and-api-requirements.md#use-runtime-checks-in-your-javascript-code)します。</span><span class="sxs-lookup"><span data-stu-id="1d9eb-122">For more information and examples of using `isSetSupported`, see [Specify Office applications and API requirements](../develop/specify-office-hosts-and-api-requirements.md), especially [Use runtime checks in your JavaScript code](../develop/specify-office-hosts-and-api-requirements.md#use-runtime-checks-in-your-javascript-code).</span></span> <span data-ttu-id="1d9eb-123">(この記事の [マニフェストの要件要素を設定](../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest) するセクションは、リボン1.1 には適用されません。)</span><span class="sxs-lookup"><span data-stu-id="1d9eb-123">(The section [Set the Requirements element in the manifest](../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest) of that article does not apply to Ribbon 1.1.)</span></span>

## <a name="shared-runtime-required"></a><span data-ttu-id="1d9eb-124">共有ランタイムが必要</span><span class="sxs-lookup"><span data-stu-id="1d9eb-124">Shared runtime required</span></span>

<span data-ttu-id="1d9eb-125">この記事で説明されている API とマニフェストのマークアップでは、アドインのマニフェストで共有ランタイムを使用するよう指定されている必要があります。</span><span class="sxs-lookup"><span data-stu-id="1d9eb-125">The APIs and manifest markup described in this article require that the add-in's manifest specify that it should use a shared runtime.</span></span> <span data-ttu-id="1d9eb-126">これを行うには、次の手順を実行します。</span><span class="sxs-lookup"><span data-stu-id="1d9eb-126">To do this take the following steps.</span></span>

1. <span data-ttu-id="1d9eb-127">マニフェストの [Runtimes](../reference/manifest/runtimes.md) 要素で、子要素の `<Runtime resid="Contoso.SharedRuntime.Url" lifetime="long" />` を追加します。</span><span class="sxs-lookup"><span data-stu-id="1d9eb-127">In the [Runtimes](../reference/manifest/runtimes.md) element in the manifest, add the following child element: `<Runtime resid="Contoso.SharedRuntime.Url" lifetime="long" />`.</span></span> <span data-ttu-id="1d9eb-128">(マニフェストに `<Runtimes>` 要素がまだない場合は、`VersionOverrides` セクションの `<Host>` 要素の下に最初の子要素として作成します。)</span><span class="sxs-lookup"><span data-stu-id="1d9eb-128">(If there isn't already a `<Runtimes>` element in the manifest, create it as the first child under the `<Host>` element in the `VersionOverrides` section.)</span></span>
2. <span data-ttu-id="1d9eb-129">マニフェストの [Resources.Urls](../reference/manifest/resources.md) セクションで、子要素の `<bt:Url id="Contoso.SharedRuntime.Url" DefaultValue="https://{MyDomain}/{path-to-start-page}" />` を追加します。ここでは、`{MyDomain}` はアドインのドメインで、`{path-to-start-page}` はアドインの開始ページのパスになります (例: `<bt:Url id="Contoso.SharedRuntime.Url" DefaultValue="https://localhost:3000/index.html" />`)。</span><span class="sxs-lookup"><span data-stu-id="1d9eb-129">In the [Resources.Urls](../reference/manifest/resources.md) section of the manifest, add the following child element: `<bt:Url id="Contoso.SharedRuntime.Url" DefaultValue="https://{MyDomain}/{path-to-start-page}" />`, where `{MyDomain}` is the domain of the add-in and `{path-to-start-page}` is the path for the start page of the add-in; for example: `<bt:Url id="Contoso.SharedRuntime.Url" DefaultValue="https://localhost:3000/index.html" />`.</span></span>
3. <span data-ttu-id="1d9eb-130">アドインに作業ウィンドウ、関数ファイル、あるいは Excel のカスタム関数が含まれているかどうかに応じて、次の 3 つの中から 1 つまたは複数の手順を実行する必要があります。</span><span class="sxs-lookup"><span data-stu-id="1d9eb-130">Depending on whether your add-in contains a task pane, a function file, or an Excel custom function, you must do one or more of the following three steps:</span></span>

    - <span data-ttu-id="1d9eb-131">アドインに作業ウィンドウが含まれている場合は、[Action](../reference/manifest/action.md).[SourceLocation](../reference/manifest/sourcelocation.md) 要素の `resid` 属性を、手順 1 で `<Runtime>` 要素の `resid` に使用したのとまったく同じ文字列に設定します。たとえば、`Contoso.SharedRuntime.Url` のようにします。</span><span class="sxs-lookup"><span data-stu-id="1d9eb-131">If the add-in contains a task pane, set the `resid` attribute of the [Action](../reference/manifest/action.md).[SourceLocation](../reference/manifest/sourcelocation.md) element to exactly the same string as you used for the `resid` of the `<Runtime>` element in step 1; for example, `Contoso.SharedRuntime.Url`.</span></span> <span data-ttu-id="1d9eb-132">そうすると要素は `<SourceLocation resid="Contoso.SharedRuntime.Url"/>` のようになります。</span><span class="sxs-lookup"><span data-stu-id="1d9eb-132">The element should look like this: `<SourceLocation resid="Contoso.SharedRuntime.Url"/>`.</span></span>
    - <span data-ttu-id="1d9eb-133">アドインに Excel カスタム関数が含まれている場合は、[Page](../reference/manifest/page.md).[SourceLocation](../reference/manifest/sourcelocation.md) 要素の `resid` 属性を、手順 1 で`<Runtime>` 要素の `resid` に使用したのとまったく同じ文字列に設定します。たとえば、`Contoso.SharedRuntime.Url` のようにします。</span><span class="sxs-lookup"><span data-stu-id="1d9eb-133">If the add-in contains an Excel custom function, set the `resid` attribute of the [Page](../reference/manifest/page.md).[SourceLocation](../reference/manifest/sourcelocation.md) element exactly the same string as you used for the `resid` of the `<Runtime>` element in step 1; for example, `Contoso.SharedRuntime.Url`.</span></span> <span data-ttu-id="1d9eb-134">そうすると要素は `<SourceLocation resid="Contoso.SharedRuntime.Url"/>` のようになります。</span><span class="sxs-lookup"><span data-stu-id="1d9eb-134">The element should look like this: `<SourceLocation resid="Contoso.SharedRuntime.Url"/>`.</span></span>
    - <span data-ttu-id="1d9eb-135">アドインに関数ファイルが含まれている場合は、[FunctionFile](../reference/manifest/functionfile.md) 要素の `resid` 属性を、手順 1 で `<Runtime>` 要素の `resid` に使用したのとまったく同じ文字列に設定します。たとえば、`Contoso.SharedRuntime.Url` のようにします。</span><span class="sxs-lookup"><span data-stu-id="1d9eb-135">If the add-in contains a function file, set the `resid` attribute of the [FunctionFile](../reference/manifest/functionfile.md) element to exactly the same string as you used for the `resid` of the `<Runtime>` element in step 1; for example, `Contoso.SharedRuntime.Url`.</span></span> <span data-ttu-id="1d9eb-136">そうすると要素は `<FunctionFile resid="Contoso.SharedRuntime.Url"/>` のようになります。</span><span class="sxs-lookup"><span data-stu-id="1d9eb-136">The element should look like this: `<FunctionFile resid="Contoso.SharedRuntime.Url"/>`.</span></span>

## <a name="set-the-default-state-to-disabled"></a><span data-ttu-id="1d9eb-137">既定の状態を無効に設定する</span><span class="sxs-lookup"><span data-stu-id="1d9eb-137">Set the default state to disabled</span></span>

<span data-ttu-id="1d9eb-138">既定では、Office アプリケーションの起動時にすべてのアドイン コマンドが有効になります。</span><span class="sxs-lookup"><span data-stu-id="1d9eb-138">By default, any Add-in Command is enabled when the Office application launches.</span></span> <span data-ttu-id="1d9eb-139">Office アプリケーションの起動時にカスタム ボタンまたはメニュー項目を無効にするには、マニフェストで指定します。</span><span class="sxs-lookup"><span data-stu-id="1d9eb-139">If you want a custom button or menu item to be disabled when the Office application launches, you specify this in the manifest.</span></span> <span data-ttu-id="1d9eb-140">コントロールの宣言の [Action](../reference/manifest/action.md) 要素の*直下* (内部ではない) に、[Enabled](../reference/manifest/enabled.md) 要素 (値は `false`) を追加するだけで無効にすることができます。</span><span class="sxs-lookup"><span data-stu-id="1d9eb-140">Just add an [Enabled](../reference/manifest/enabled.md) element (with the value `false`) immediately *below* (not inside) the [Action](../reference/manifest/action.md) element in the declaration of the control.</span></span> <span data-ttu-id="1d9eb-141">基本的な構造を次に示します。</span><span class="sxs-lookup"><span data-stu-id="1d9eb-141">The following shows the basic structure:</span></span>

```xml
<OfficeApp ...>
  ...
  <VersionOverrides ...>
    ...
    <Hosts>
      <Host ...>
        ...
        <DesktopFormFactor>
          <ExtensionPoint ...>
            <CustomTab ...>
              ...
              <Group ...>
                ...
                <Control ... id="MyButton">
                  ...
                  <Action ...>
                  <Enabled>false</Enabled>
...
</OfficeApp>
```

## <a name="change-the-state-programmatically"></a><span data-ttu-id="1d9eb-142">プログラムで状態を変更する</span><span class="sxs-lookup"><span data-stu-id="1d9eb-142">Change the state programmatically</span></span>

<span data-ttu-id="1d9eb-143">アドイン コマンドの有効な状態を変更するには、以下の手順が重要になります。</span><span class="sxs-lookup"><span data-stu-id="1d9eb-143">The essential steps to changing the enabled status of an Add-in Command are:</span></span>

1. <span data-ttu-id="1d9eb-144">マニフェストで指定された ID でコマンドとその親タブを指定する [RibbonUpdaterData](/javascript/api/office/office.ribbonupdaterdata) オブジェクトを作成し、コマンドの状態を有効か無効かに指定します。</span><span class="sxs-lookup"><span data-stu-id="1d9eb-144">Create a [RibbonUpdaterData](/javascript/api/office/office.ribbonupdaterdata) object that (1) specifies the command, and its parent tab, by their IDs as specified in the manifest; and (2) specifies the enabled or disabled state of the command.</span></span>
2. <span data-ttu-id="1d9eb-145">**RibbonUpdaterData** オブジェクトを [Office.ribbon.requestUpdate()](/javascript/api/office/office.ribbon?view=common-js#requestupdate-input-) メソッドに渡します。</span><span class="sxs-lookup"><span data-stu-id="1d9eb-145">Pass the **RibbonUpdaterData** object to the [Office.ribbon.requestUpdate()](/javascript/api/office/office.ribbon?view=common-js#requestupdate-input-) method.</span></span>

<span data-ttu-id="1d9eb-146">次に簡単な例を示します。</span><span class="sxs-lookup"><span data-stu-id="1d9eb-146">The following is a simple example.</span></span> <span data-ttu-id="1d9eb-147">"MyButton" と "OfficeAddinTab1" はマニフェストからコピーされます。</span><span class="sxs-lookup"><span data-stu-id="1d9eb-147">Note that "MyButton" and "OfficeAddinTab1" are copied from the manifest.</span></span>

```javascript
function enableButton() {
    Office.ribbon.requestUpdate({
        tabs: [
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

<span data-ttu-id="1d9eb-148">また、**RibbonUpdateData** オブジェクトを簡単に構築できるように、いくつかのインターフェイスも (何種類か) 用意しています。</span><span class="sxs-lookup"><span data-stu-id="1d9eb-148">We also provide several interfaces (types) to make it easier to construct the **RibbonUpdateData** object.</span></span> <span data-ttu-id="1d9eb-149">以下は、TypeScript の同じ例であり、インターフェイスを使用したものです。</span><span class="sxs-lookup"><span data-stu-id="1d9eb-149">The following is the equivalent example in TypeScript and it makes use of these types.</span></span>

```typescript
const enableButton = async () => {
    const button: Control = {id: "MyButton", enabled: true};
    const parentTab: Tab = {id: "OfficeAddinTab1", controls: [button]};
    const ribbonUpdater: RibbonUpdaterData = { tabs: [parentTab]};
    await Office.ribbon.requestUpdate(ribbonUpdater);
}
```

<span data-ttu-id="1d9eb-150">Office では、リボンの状態を更新するタイミングが制御されます。</span><span class="sxs-lookup"><span data-stu-id="1d9eb-150">Office controls when it updates the state of the ribbon.</span></span> <span data-ttu-id="1d9eb-151">**requestUpdate ()** メソッドが、更新の要求をキューイングします。</span><span class="sxs-lookup"><span data-stu-id="1d9eb-151">The **requestUpdate()** method queues a request to update.</span></span> <span data-ttu-id="1d9eb-152">このメソッドによる Promise オブジェクトの解決は、リボンが実際に更新されたときではなく、要求がキューに登録された直後に行われます。</span><span class="sxs-lookup"><span data-stu-id="1d9eb-152">The method will resolve the Promise object as soon as it has queued the request, not when the ribbon actually updates.</span></span>

## <a name="change-the-state-in-response-to-an-event"></a><span data-ttu-id="1d9eb-153">イベントに応じて状態を変更する</span><span class="sxs-lookup"><span data-stu-id="1d9eb-153">Change the state in response to an event</span></span>

<span data-ttu-id="1d9eb-154">リボンの状態を変更する一般的なシナリオは、ユーザーが開始したイベントがアドインのコンテキストを変更したときです。</span><span class="sxs-lookup"><span data-stu-id="1d9eb-154">A common scenario in which the ribbon state should change is when a user-initiated event changes the add-in context.</span></span>

<span data-ttu-id="1d9eb-155">グラフがアクティブになったときにのみボタンを有効にするシナリオを考えます。</span><span class="sxs-lookup"><span data-stu-id="1d9eb-155">Consider a scenario in which a button should be enabled when, and only when, a chart is activated.</span></span> <span data-ttu-id="1d9eb-156">まず、マニフェストのボタンの [Enabled](../reference/manifest/enabled.md) 要素を `false` に設定します。</span><span class="sxs-lookup"><span data-stu-id="1d9eb-156">The first step is to set the [Enabled](../reference/manifest/enabled.md) element for the button in the manifest to `false`.</span></span> <span data-ttu-id="1d9eb-157">例については上記を参照してください。</span><span class="sxs-lookup"><span data-stu-id="1d9eb-157">See above for an example.</span></span>

<span data-ttu-id="1d9eb-158">次に、ハンドラーを割り当てます。</span><span class="sxs-lookup"><span data-stu-id="1d9eb-158">Second, assign handlers.</span></span> <span data-ttu-id="1d9eb-159">これは通常、ハンドラー (後の手順で作成) をワークシート内のすべてのグラフの **onActivated** および **onDeactivated** イベントに割り当てる以下の例のように、**Office.onReady** メソッドで行います。</span><span class="sxs-lookup"><span data-stu-id="1d9eb-159">This is commonly done in the **Office.onReady** method as in the following example which assigns handlers (created in a later step) to the **onActivated** and **onDeactivated** events of all the charts in the worksheet.</span></span>

```javascript
Office.onReady(async () => {
    await Excel.run(context => {
        var charts = context.workbook.worksheets
            .getActiveWorksheet()
            .charts;
        charts.onActivated.add(enableChartFormat);
        charts.onDeactivated.add(disableChartFormat);
        return context.sync();
    });
});
```

<span data-ttu-id="1d9eb-160">そして、`enableChartFormat` ハンドラーを定義します。</span><span class="sxs-lookup"><span data-stu-id="1d9eb-160">Third, define the `enableChartFormat` handler.</span></span> <span data-ttu-id="1d9eb-161">以下は簡単な例ですが、より信頼性の高い方法でコントロールの状態を変更する場合については、後述の「[ベスト プラクティス: コントロールの状態エラーのテスト](#best-practice-test-for-control-status-errors)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="1d9eb-161">The following is a simple example, but see [Best practice: Test for control status errors](#best-practice-test-for-control-status-errors) below for a more robust way of changing a control's status.</span></span>

```javascript
function enableChartFormat() {
    var button = {id: "ChartFormatButton", enabled: true};
    var parentTab = {id: "CustomChartTab", controls: [button]};
    var ribbonUpdater = {tabs: [parentTab]};
    await Office.ribbon.requestUpdate(ribbonUpdater);
}
```

<span data-ttu-id="1d9eb-162">最後に、`disableChartFormat` ハンドラーを定義します。</span><span class="sxs-lookup"><span data-stu-id="1d9eb-162">Fourth, define the `disableChartFormat` handler.</span></span> <span data-ttu-id="1d9eb-163">`enableChartFormat` と同じですが、ボタン オブジェクトの **enabled** プロパティを `false` に設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="1d9eb-163">It would be identical to `enableChartFormat` except that the **enabled** property of the button object would be set to `false`.</span></span>

## <a name="best-practice-test-for-control-status-errors"></a><span data-ttu-id="1d9eb-164">ベスト プラクティス: コントロールの状態エラーのテスト</span><span class="sxs-lookup"><span data-stu-id="1d9eb-164">Best practice: Test for control status errors</span></span>

<span data-ttu-id="1d9eb-165">状況によっては、`requestUpdate` が呼び出された後でもリボンが再描画されず、コントロールのクリック可能な状態が変更されない場合があります。</span><span class="sxs-lookup"><span data-stu-id="1d9eb-165">In some circumstances, the ribbon does not repaint after `requestUpdate` is called, so the control's clickable status does not change.</span></span> <span data-ttu-id="1d9eb-166">そこで、アドインのベスト プラクティスとして、コントロールの状態を追跡することが挙げられます。</span><span class="sxs-lookup"><span data-stu-id="1d9eb-166">For this reason it is a best practice for the add-in to keep track of the status of its controls.</span></span> <span data-ttu-id="1d9eb-167">アドインは次のルールに準拠する必要があります。</span><span class="sxs-lookup"><span data-stu-id="1d9eb-167">The add-in should conform to these rules:</span></span>

1. <span data-ttu-id="1d9eb-168">`requestUpdate` が呼び出された場合はいつでも、コードがカスタム ボタンとメニュー項目の意図した状態を記録する必要があります。</span><span class="sxs-lookup"><span data-stu-id="1d9eb-168">Whenever `requestUpdate` is called, the code should record the intended state of the custom buttons and menu items.</span></span>
2. <span data-ttu-id="1d9eb-169">カスタム コントロールがクリックされたら、ハンドラーの最初のコードが、ボタンがクリック可能になっているかどうかを確認する必要があります。</span><span class="sxs-lookup"><span data-stu-id="1d9eb-169">When a custom control is clicked, the first code in the handler, should check to see if the button should have been clickable.</span></span> <span data-ttu-id="1d9eb-170">クリック可能でない場合は、コードがエラーの報告または記録を行い、ボタンを意図した状態に設定し直す必要があります。</span><span class="sxs-lookup"><span data-stu-id="1d9eb-170">If shouldn't have been, the code should report or log an error and try again to set the buttons to the intended state.</span></span>

<span data-ttu-id="1d9eb-171">次の例は、ボタンを無効にし、ボタンの状態を記録する関数を示しています。</span><span class="sxs-lookup"><span data-stu-id="1d9eb-171">The following example shows a function that disables a button and records the button's status.</span></span> <span data-ttu-id="1d9eb-172">`chartFormatButtonEnabled` は、マニフェスト内のボタンの [Enabled](../reference/manifest/enabled.md) 要素と同じ値に初期化されるグローバルなブール変数です。</span><span class="sxs-lookup"><span data-stu-id="1d9eb-172">Note that `chartFormatButtonEnabled` is a global boolean variable that is initialized to the same value as the [Enabled](../reference/manifest/enabled.md) element for the button in the manifest.</span></span>

```javascript
function disableChartFormat() {
    var button = {id: "ChartFormatButton", enabled: false};
    var parentTab = {id: "CustomChartTab", controls: [button]};
    var ribbonUpdater = {tabs: [parentTab]};
    await Office.ribbon.requestUpdate(ribbonUpdater);

    chartFormatButtonEnabled = false;
}
```

<span data-ttu-id="1d9eb-173">次の例は、ボタンのハンドラーがボタンの不正な状態をテストする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="1d9eb-173">The following example shows how the button's handler tests for an incorrect state of the button.</span></span> <span data-ttu-id="1d9eb-174">`reportError` は、エラーを表示または記録する関数です。</span><span class="sxs-lookup"><span data-stu-id="1d9eb-174">Note that `reportError` is a function that shows or logs an error.</span></span>

```javascript
function chartFormatButtonHandler() {
    if (chartFormatButtonEnabled) {

        // Do work here

    } else {
        // Report the error and try again to disable.
        reportError("That action is not possible at this time.");
        disableChartFormat();
    }
}
```

## <a name="error-handling"></a><span data-ttu-id="1d9eb-175">エラー処理</span><span class="sxs-lookup"><span data-stu-id="1d9eb-175">Error handling</span></span>

<span data-ttu-id="1d9eb-176">一部のシナリオでは、Office はリボンを更新できず、エラーを返します。</span><span class="sxs-lookup"><span data-stu-id="1d9eb-176">In some scenarios, Office is unable to update the ribbon and will return an error.</span></span> <span data-ttu-id="1d9eb-177">たとえば、アドインがアップグレードされ、アップグレードされたアドインに異なるカスタム アドイン コマンドのセットがある場合は、Office アプリケーションを閉じてから、もう一度開く必要があります。</span><span class="sxs-lookup"><span data-stu-id="1d9eb-177">For example, if the add-in is upgraded and the upgraded add-in has a different set of custom add-in commands, then the Office application must be closed and reopened.</span></span> <span data-ttu-id="1d9eb-178">それまでの間、`requestUpdate` メソッドは `HostRestartNeeded` エラーを返します。</span><span class="sxs-lookup"><span data-stu-id="1d9eb-178">Until it is, the `requestUpdate` method will return the error `HostRestartNeeded`.</span></span> <span data-ttu-id="1d9eb-179">このエラーの処理方法の例を次に示します。</span><span class="sxs-lookup"><span data-stu-id="1d9eb-179">The following is an example of how to handle this error.</span></span> <span data-ttu-id="1d9eb-180">この場合、`reportError` メソッドがユーザーにエラーを表示します。</span><span class="sxs-lookup"><span data-stu-id="1d9eb-180">In this case, the `reportError` method displays the error to the user.</span></span>

```javascript
function disableChartFormat() {
    try {
        var button = {id: "ChartFormatButton", enabled: false};
        var parentTab = {id: "CustomChartTab", controls: [button]};
        var ribbonUpdater = {tabs: [parentTab]};
        await Office.ribbon.requestUpdate(ribbonUpdater);

        chartFormatButtonEnabled = false;
    }
    catch(error) {
        if (error.code == "HostRestartNeeded"){
            reportError("Contoso Awesome Add-in has been upgraded. Please save your work, close the Office application, and restart it.");
        }
    }
}
```

## <a name="test-for-platform-support-with-requirement-sets"></a><span data-ttu-id="1d9eb-181">要件セットを使用したプラットフォーム サポートのテスト</span><span class="sxs-lookup"><span data-stu-id="1d9eb-181">Test for platform support with requirement sets</span></span>

<span data-ttu-id="1d9eb-182">要件セットは、API メンバーの名前付きグループです。</span><span class="sxs-lookup"><span data-stu-id="1d9eb-182">Requirement sets are named groups of API members.</span></span> <span data-ttu-id="1d9eb-183">Office アドインは、マニフェストで指定されている要件セットを使用するか、ランタイムチェックを使用して、Office アプリケーションがアドインに必要な Api をサポートしているかどうかを判断します。</span><span class="sxs-lookup"><span data-stu-id="1d9eb-183">Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office application supports APIs that an add-in needs.</span></span> <span data-ttu-id="1d9eb-184">詳細については、「 [Office のバージョンと要件セット](../develop/office-versions-and-requirement-sets.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="1d9eb-184">For more information, see [Office versions and requirement sets](../develop/office-versions-and-requirement-sets.md).</span></span>

<span data-ttu-id="1d9eb-185">API を有効化/無効化するには、次の要件セットをサポートしている必要があります。</span><span class="sxs-lookup"><span data-stu-id="1d9eb-185">The enable/disable APIs require support of the following requirement set:</span></span>

- [<span data-ttu-id="1d9eb-186">RibbonApi 1.1</span><span class="sxs-lookup"><span data-stu-id="1d9eb-186">RibbonApi 1.1</span></span>](../reference/requirement-sets/ribbon-api-requirement-sets.md)

