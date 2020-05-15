---
title: アドイン コマンドを有効または無効にする
description: Office Web アドインのカスタム リボン ボタンとメニュー項目の有効または無効の状態を変更する方法について説明します。
ms.date: 05/11/2020
localization_priority: Priority
ms.openlocfilehash: fa4830c0112486bbad7a13edf78e0c8c4277e143
ms.sourcegitcommit: 682d18c9149b1153f9c38d28e2a90384e6a261dc
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/13/2020
ms.locfileid: "44217894"
---
# <a name="enable-and-disable-add-in-commands"></a><span data-ttu-id="4ca29-103">アドイン コマンドを有効または無効にする</span><span class="sxs-lookup"><span data-stu-id="4ca29-103">Enable and Disable Add-in Commands</span></span>

<span data-ttu-id="4ca29-104">アドインの一部の機能を特定のコンテキストでのみ使用可能にする必要がある場合、カスタム アドイン コマンドをプログラムで有効または無効にすることができます。</span><span class="sxs-lookup"><span data-stu-id="4ca29-104">When some functionality in your add-in should only be available in certain contexts, you can programmatically enable or disable your custom Add-in Commands.</span></span> <span data-ttu-id="4ca29-105">たとえば、表の見出しを変更する関数は、カーソルが表の中にある場合にのみ有効にする必要があります。</span><span class="sxs-lookup"><span data-stu-id="4ca29-105">For example, a function that changes the header of a table should only be enabled when the cursor is in a table.</span></span>

<span data-ttu-id="4ca29-106">また、Office のホスト アプリケーションを開いたときにコマンドを有効にするか無効にするかを指定することもできます。</span><span class="sxs-lookup"><span data-stu-id="4ca29-106">You can also specify whether the command is enabled or disabled when the Office host application opens.</span></span>

> [!NOTE]
> <span data-ttu-id="4ca29-107">この記事は、以下のドキュメントについて既に理解していることを前提としています。</span><span class="sxs-lookup"><span data-stu-id="4ca29-107">This article assumes that you are familiar with the following documentation.</span></span> <span data-ttu-id="4ca29-108">最近、アドイン コマンド (カスタム メニュー項目とリボン ボタン) を使用してない場合は、ドキュメントをご確認ください。</span><span class="sxs-lookup"><span data-stu-id="4ca29-108">Please review it if you haven't worked with Add-in Commands (custom menu items and ribbon buttons) recently.</span></span>
>
> [<span data-ttu-id="4ca29-109">アドイン コマンドの基本概念</span><span class="sxs-lookup"><span data-stu-id="4ca29-109">Basic concepts for Add-in Commands</span></span>](add-in-commands.md)

## <a name="rules-and-gotchas"></a><span data-ttu-id="4ca29-110">ルールと注意事項</span><span class="sxs-lookup"><span data-stu-id="4ca29-110">Rules and gotchas</span></span>

### <a name="single-line-ribbon-in-office-on-the-web"></a><span data-ttu-id="4ca29-111">Office on the web の単一行のリボン</span><span class="sxs-lookup"><span data-stu-id="4ca29-111">Single-line ribbon in Office on the web</span></span>

<span data-ttu-id="4ca29-112">この記事で説明されている API と、マニフェストのマークアップは、Office on the web では単一行のリボンにのみ影響します。</span><span class="sxs-lookup"><span data-stu-id="4ca29-112">In Office on the web, the APIs and manifest markup described in this article only affect the single-line ribbon.</span></span> <span data-ttu-id="4ca29-113">複数行のリボンには影響しません。</span><span class="sxs-lookup"><span data-stu-id="4ca29-113">They have no effect on the multiline ribbon.</span></span> <span data-ttu-id="4ca29-114">デスクトップ Office では両方のリボンに影響します。</span><span class="sxs-lookup"><span data-stu-id="4ca29-114">They affect both ribbons for desktop Office.</span></span> <span data-ttu-id="4ca29-115">2 つのリボンの詳細については、「[シンプル リボンを使用する](https://support.office.com/article/Use-the-Simplified-Ribbon-44bef9c3-295d-4092-b7f0-f471fa629a98)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="4ca29-115">For more information about the two ribbons, see [Use the simplified ribbon](https://support.office.com/article/Use-the-Simplified-Ribbon-44bef9c3-295d-4092-b7f0-f471fa629a98).</span></span>

### <a name="shared-runtime-required"></a><span data-ttu-id="4ca29-116">共有ランタイムが必要</span><span class="sxs-lookup"><span data-stu-id="4ca29-116">Shared runtime required</span></span>

<span data-ttu-id="4ca29-117">この記事で説明されている API とマニフェストのマークアップでは、アドインのマニフェストで共有ランタイムを使用するよう指定されている必要があります。</span><span class="sxs-lookup"><span data-stu-id="4ca29-117">The APIs and manifest markup described in this article require that the add-in's manifest specify that it should use a shared runtime.</span></span> <span data-ttu-id="4ca29-118">これを行うには、次の手順を実行します。</span><span class="sxs-lookup"><span data-stu-id="4ca29-118">To do this take the following steps.</span></span>

1. <span data-ttu-id="4ca29-119">マニフェストの [Runtimes](../reference/manifest/runtimes.md) 要素で、子要素の `<Runtime resid="Contoso.SharedRuntime.Url" lifetime="long" />` を追加します。</span><span class="sxs-lookup"><span data-stu-id="4ca29-119">In the [Runtimes](../reference/manifest/runtimes.md) element in the manifest, add the following child element: `<Runtime resid="Contoso.SharedRuntime.Url" lifetime="long" />`.</span></span> <span data-ttu-id="4ca29-120">(マニフェストに `<Runtimes>` 要素がまだない場合は、`VersionOverrides` セクションの `<Host>` 要素の下に最初の子要素として作成します。)</span><span class="sxs-lookup"><span data-stu-id="4ca29-120">(If there isn't already a `<Runtimes>` element in the manifest, create it as the first child under the `<Host>` element in the `VersionOverrides` section.)</span></span>
2. <span data-ttu-id="4ca29-121">マニフェストの [Resources.Urls](../reference/manifest/resources.md) セクションで、子要素の `<bt:Url id="Contoso.SharedRuntime.Url" DefaultValue="https://{MyDomain}/{path-to-start-page}" />` を追加します。ここでは、`{MyDomain}` はアドインのドメインで、`{path-to-start-page}` はアドインの開始ページのパスになります (例: `<bt:Url id="Contoso.SharedRuntime.Url" DefaultValue="https://localhost:3000/index.html" />`)。</span><span class="sxs-lookup"><span data-stu-id="4ca29-121">In the [Resources.Urls](../reference/manifest/resources.md) section of the manifest, add the following child element: `<bt:Url id="Contoso.SharedRuntime.Url" DefaultValue="https://{MyDomain}/{path-to-start-page}" />`, where `{MyDomain}` is the domain of the add-in and `{path-to-start-page}` is the path for the start page of the add-in; for example: `<bt:Url id="Contoso.SharedRuntime.Url" DefaultValue="https://localhost:3000/index.html" />`.</span></span>
3. <span data-ttu-id="4ca29-122">アドインに作業ウィンドウ、関数ファイル、あるいは Excel のカスタム関数が含まれているかどうかに応じて、次の 3 つの中から 1 つまたは複数の手順を実行する必要があります。</span><span class="sxs-lookup"><span data-stu-id="4ca29-122">Depending on whether your add-in contains a task pane, a function file, or an Excel custom function, you must do one or more of the following three steps:</span></span>

    - <span data-ttu-id="4ca29-123">アドインに作業ウィンドウが含まれている場合は、[Action](../reference/manifest/action.md).[SourceLocation](../reference/manifest/sourcelocation.md) 要素の `resid` 属性を、手順 1 で `<Runtime>` 要素の `resid` に使用したのとまったく同じ文字列に設定します。たとえば、`Contoso.SharedRuntime.Url` のようにします。</span><span class="sxs-lookup"><span data-stu-id="4ca29-123">If the add-in contains a task pane, set the `resid` attribute of the [Action](../reference/manifest/action.md).[SourceLocation](../reference/manifest/sourcelocation.md) element to exactly the same string as you used for the `resid` of the `<Runtime>` element in step 1; for example, `Contoso.SharedRuntime.Url`.</span></span> <span data-ttu-id="4ca29-124">そうすると要素は `<SourceLocation resid="Contoso.SharedRuntime.Url"/>` のようになります。</span><span class="sxs-lookup"><span data-stu-id="4ca29-124">The element should look like this: `<SourceLocation resid="Contoso.SharedRuntime.Url"/>`.</span></span>
    - <span data-ttu-id="4ca29-125">アドインに Excel カスタム関数が含まれている場合は、[Page](../reference/manifest/page.md).[SourceLocation](../reference/manifest/sourcelocation.md) 要素の `resid` 属性を、手順 1 で`<Runtime>` 要素の `resid` に使用したのとまったく同じ文字列に設定します。たとえば、`Contoso.SharedRuntime.Url` のようにします。</span><span class="sxs-lookup"><span data-stu-id="4ca29-125">If the add-in contains an Excel custom function, set the `resid` attribute of the [Page](../reference/manifest/page.md).[SourceLocation](../reference/manifest/sourcelocation.md) element exactly the same string as you used for the `resid` of the `<Runtime>` element in step 1; for example, `Contoso.SharedRuntime.Url`.</span></span> <span data-ttu-id="4ca29-126">そうすると要素は `<SourceLocation resid="Contoso.SharedRuntime.Url"/>` のようになります。</span><span class="sxs-lookup"><span data-stu-id="4ca29-126">The element should look like this: `<SourceLocation resid="Contoso.SharedRuntime.Url"/>`.</span></span>
    - <span data-ttu-id="4ca29-127">アドインに関数ファイルが含まれている場合は、[FunctionFile](../reference/manifest/functionfile.md) 要素の `resid` 属性を、手順 1 で `<Runtime>` 要素の `resid` に使用したのとまったく同じ文字列に設定します。たとえば、`Contoso.SharedRuntime.Url` のようにします。</span><span class="sxs-lookup"><span data-stu-id="4ca29-127">If the add-in contains a function file, set the `resid` attribute of the [FunctionFile](../reference/manifest/functionfile.md) element to exactly the same string as you used for the `resid` of the `<Runtime>` element in step 1; for example, `Contoso.SharedRuntime.Url`.</span></span> <span data-ttu-id="4ca29-128">そうすると要素は `<FunctionFile resid="Contoso.SharedRuntime.Url"/>` のようになります。</span><span class="sxs-lookup"><span data-stu-id="4ca29-128">The element should look like this: `<FunctionFile resid="Contoso.SharedRuntime.Url"/>`.</span></span>

## <a name="set-the-default-state-to-disabled"></a><span data-ttu-id="4ca29-129">既定の状態を無効に設定する</span><span class="sxs-lookup"><span data-stu-id="4ca29-129">Set the default state to disabled</span></span>

<span data-ttu-id="4ca29-130">既定では、Office アプリケーションの起動時にすべてのアドイン コマンドが有効になります。</span><span class="sxs-lookup"><span data-stu-id="4ca29-130">By default, any Add-in Command is enabled when the Office application launches.</span></span> <span data-ttu-id="4ca29-131">Office アプリケーションの起動時にカスタム ボタンまたはメニュー項目を無効にするには、マニフェストで指定します。</span><span class="sxs-lookup"><span data-stu-id="4ca29-131">If you want a custom button or menu item to be disabled when the Office application launches, you specify this in the manifest.</span></span> <span data-ttu-id="4ca29-132">コントロールの宣言の [Action](../reference/manifest/action.md) 要素の*直下* (内部ではない) に、[Enabled](../reference/manifest/enabled.md) 要素 (値は `false`) を追加するだけで無効にすることができます。</span><span class="sxs-lookup"><span data-stu-id="4ca29-132">Just add an [Enabled](../reference/manifest/enabled.md) element (with the value `false`) immediately *below* (not inside) the [Action](../reference/manifest/action.md) element in the declaration of the control.</span></span> <span data-ttu-id="4ca29-133">基本的な構造を次に示します。</span><span class="sxs-lookup"><span data-stu-id="4ca29-133">The following shows the basic structure:</span></span>

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

## <a name="change-the-state-programmatically"></a><span data-ttu-id="4ca29-134">プログラムで状態を変更する</span><span class="sxs-lookup"><span data-stu-id="4ca29-134">Change the state programmatically</span></span>

<span data-ttu-id="4ca29-135">アドイン コマンドの有効な状態を変更するには、以下の手順が重要になります。</span><span class="sxs-lookup"><span data-stu-id="4ca29-135">The essential steps to changing the enabled status of an Add-in Command are:</span></span>

1. <span data-ttu-id="4ca29-136">マニフェストで指定された ID でコマンドとその親タブを指定する [RibbonUpdaterData](/javascript/api/office/office.ribbonupdaterdata) オブジェクトを作成し、コマンドの状態を有効か無効かに指定します。</span><span class="sxs-lookup"><span data-stu-id="4ca29-136">Create a [RibbonUpdaterData](/javascript/api/office/office.ribbonupdaterdata) object that (1) specifies the command, and its parent tab, by their IDs as specified in the manifest; and (2) specifies the enabled or disabled state of the command.</span></span>
2. <span data-ttu-id="4ca29-137">**RibbonUpdaterData** オブジェクトを [Office.ribbon.requestUpdate()](/javascript/api/office/office.ribbon?view=common-js#requestupdate-input-) メソッドに渡します。</span><span class="sxs-lookup"><span data-stu-id="4ca29-137">Pass the **RibbonUpdaterData** object to the [Office.ribbon.requestUpdate()](/javascript/api/office/office.ribbon?view=common-js#requestupdate-input-) method.</span></span>

<span data-ttu-id="4ca29-138">次に簡単な例を示します。</span><span class="sxs-lookup"><span data-stu-id="4ca29-138">The following is a simple example.</span></span> <span data-ttu-id="4ca29-139">"MyButton" と "OfficeAddinTab1" はマニフェストからコピーされます。</span><span class="sxs-lookup"><span data-stu-id="4ca29-139">Note that "MyButton" and "OfficeAddinTab1" are copied from the manifest.</span></span>

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

<span data-ttu-id="4ca29-140">また、**RibbonUpdateData** オブジェクトを簡単に構築できるように、いくつかのインターフェイスも (何種類か) 用意しています。</span><span class="sxs-lookup"><span data-stu-id="4ca29-140">We also provide several interfaces (types) to make it easier to construct the **RibbonUpdateData** object.</span></span> <span data-ttu-id="4ca29-141">以下は、TypeScript の同じ例であり、インターフェイスを使用したものです。</span><span class="sxs-lookup"><span data-stu-id="4ca29-141">The following is the equivalent example in TypeScript and it makes use of these types.</span></span>

```typescript
const enableButton = async () => {
    const button: Control = {id: "MyButton", enabled: true};
    const parentTab: Tab = {id: "OfficeAddinTab1", controls: [button]};
    const ribbonUpdater: RibbonUpdaterData = { tabs: [parentTab]};
    await Office.ribbon.requestUpdate(ribbonUpdater);
}
```

<span data-ttu-id="4ca29-142">Office では、リボンの状態を更新するタイミングが制御されます。</span><span class="sxs-lookup"><span data-stu-id="4ca29-142">Office controls when it updates the state of the ribbon.</span></span> <span data-ttu-id="4ca29-143">**requestUpdate ()** メソッドが、更新の要求をキューイングします。</span><span class="sxs-lookup"><span data-stu-id="4ca29-143">The **requestUpdate()** method queues a request to update.</span></span> <span data-ttu-id="4ca29-144">このメソッドによる Promise オブジェクトの解決は、リボンが実際に更新されたときではなく、要求がキューに登録された直後に行われます。</span><span class="sxs-lookup"><span data-stu-id="4ca29-144">The method will resolve the Promise object as soon as it has queued the request, not when the ribbon actually updates.</span></span>

## <a name="change-the-state-in-response-to-an-event"></a><span data-ttu-id="4ca29-145">イベントに応じて状態を変更する</span><span class="sxs-lookup"><span data-stu-id="4ca29-145">Change the state in response to an event</span></span>

<span data-ttu-id="4ca29-146">リボンの状態を変更する一般的なシナリオは、ユーザーが開始したイベントがアドインのコンテキストを変更したときです。</span><span class="sxs-lookup"><span data-stu-id="4ca29-146">A common scenario in which the ribbon state should change is when a user-initiated event changes the add-in context.</span></span>

<span data-ttu-id="4ca29-147">グラフがアクティブになったときにのみボタンを有効にするシナリオを考えます。</span><span class="sxs-lookup"><span data-stu-id="4ca29-147">Consider a scenario in which a button should be enabled when, and only when, a chart is activated.</span></span> <span data-ttu-id="4ca29-148">まず、マニフェストのボタンの [Enabled](../reference/manifest/enabled.md) 要素を `false` に設定します。</span><span class="sxs-lookup"><span data-stu-id="4ca29-148">The first step is to set the [Enabled](../reference/manifest/enabled.md) element for the button in the manifest to `false`.</span></span> <span data-ttu-id="4ca29-149">例については上記を参照してください。</span><span class="sxs-lookup"><span data-stu-id="4ca29-149">See above for an example.</span></span>

<span data-ttu-id="4ca29-150">次に、ハンドラーを割り当てます。</span><span class="sxs-lookup"><span data-stu-id="4ca29-150">Second, assign handlers.</span></span> <span data-ttu-id="4ca29-151">これは通常、ハンドラー (後の手順で作成) をワークシート内のすべてのグラフの **onActivated** および **onDeactivated** イベントに割り当てる以下の例のように、**Office.onReady** メソッドで行います。</span><span class="sxs-lookup"><span data-stu-id="4ca29-151">This is commonly done in the **Office.onReady** method as in the following example which assigns handlers (created in a later step) to the **onActivated** and **onDeactivated** events of all the charts in the worksheet.</span></span>

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

<span data-ttu-id="4ca29-152">そして、`enableChartFormat` ハンドラーを定義します。</span><span class="sxs-lookup"><span data-stu-id="4ca29-152">Third, define the `enableChartFormat` handler.</span></span> <span data-ttu-id="4ca29-153">以下は簡単な例ですが、より信頼性の高い方法でコントロールの状態を変更する場合については、後述の「**ベスト プラクティス: コントロールの状態エラーのテスト**」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="4ca29-153">The following is a simple example, but see **Best practice: Test for control status errors** below for a more robust way of changing a control's status.</span></span>

```javascript
function enableChartFormat() {
    var button = {id: "ChartFormatButton", enabled: true};
    var parentTab = {id: "CustomChartTab", controls: [button]};
    var ribbonUpdater = {tabs: [parentTab]};
    await Office.ribbon.requestUpdate(ribbonUpdater);
}
```

<span data-ttu-id="4ca29-154">最後に、`disableChartFormat` ハンドラーを定義します。</span><span class="sxs-lookup"><span data-stu-id="4ca29-154">Fourth, define the `disableChartFormat` handler.</span></span> <span data-ttu-id="4ca29-155">`enableChartFormat` と同じですが、ボタン オブジェクトの **enabled** プロパティを `false` に設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="4ca29-155">It would be identical to `enableChartFormat` except that the **enabled** property of the button object would be set to `false`.</span></span>

## <a name="best-practice-test-for-control-status-errors"></a><span data-ttu-id="4ca29-156">ベスト プラクティス: コントロールの状態エラーのテスト</span><span class="sxs-lookup"><span data-stu-id="4ca29-156">Best practice: Test for control status errors</span></span>

<span data-ttu-id="4ca29-157">状況によっては、`requestUpdate` が呼び出された後でもリボンが再描画されず、コントロールのクリック可能な状態が変更されない場合があります。</span><span class="sxs-lookup"><span data-stu-id="4ca29-157">In some circumstances, the ribbon does not repaint after `requestUpdate` is called, so the control's clickable status does not change.</span></span> <span data-ttu-id="4ca29-158">そこで、アドインのベスト プラクティスとして、コントロールの状態を追跡することが挙げられます。</span><span class="sxs-lookup"><span data-stu-id="4ca29-158">For this reason it is a best practice for the add-in to keep track of the status of its controls.</span></span> <span data-ttu-id="4ca29-159">アドインは次のルールに準拠する必要があります。</span><span class="sxs-lookup"><span data-stu-id="4ca29-159">The add-in should conform to these rules:</span></span>

1. <span data-ttu-id="4ca29-160">`requestUpdate` が呼び出された場合はいつでも、コードがカスタム ボタンとメニュー項目の意図した状態を記録する必要があります。</span><span class="sxs-lookup"><span data-stu-id="4ca29-160">Whenever `requestUpdate` is called, the code should record the intended state of the custom buttons and menu items.</span></span>
2. <span data-ttu-id="4ca29-161">カスタム コントロールがクリックされたら、ハンドラーの最初のコードが、ボタンがクリック可能になっているかどうかを確認する必要があります。</span><span class="sxs-lookup"><span data-stu-id="4ca29-161">When a custom control is clicked, the first code in the handler, should check to see if the button should have been clickable.</span></span> <span data-ttu-id="4ca29-162">クリック可能でない場合は、コードがエラーの報告または記録を行い、ボタンを意図した状態に設定し直す必要があります。</span><span class="sxs-lookup"><span data-stu-id="4ca29-162">If shouldn't have been, the code should report or log an error and try again to set the buttons to the intended state.</span></span>

<span data-ttu-id="4ca29-163">次の例は、ボタンを無効にし、ボタンの状態を記録する関数を示しています。</span><span class="sxs-lookup"><span data-stu-id="4ca29-163">The following example shows a function that disables a button and records the button's status.</span></span> <span data-ttu-id="4ca29-164">`chartFormatButtonEnabled` は、マニフェスト内のボタンの [Enabled](../reference/manifest/enabled.md) 要素と同じ値に初期化されるグローバルなブール変数です。</span><span class="sxs-lookup"><span data-stu-id="4ca29-164">Note that `chartFormatButtonEnabled` is a global boolean variable that is initialized to the same value as the [Enabled](../reference/manifest/enabled.md) element for the button in the manifest.</span></span>

```javascript
function disableChartFormat() {
    var button = {id: "ChartFormatButton", enabled: false};
    var parentTab = {id: "CustomChartTab", controls: [button]};
    var ribbonUpdater = {tabs: [parentTab]};
    await Office.ribbon.requestUpdate(ribbonUpdater);

    chartFormatButtonEnabled = false;
}
```

<span data-ttu-id="4ca29-165">次の例は、ボタンのハンドラーがボタンの不正な状態をテストする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="4ca29-165">The following example shows how the button's handler tests for an incorrect state of the button.</span></span> <span data-ttu-id="4ca29-166">`reportError` は、エラーを表示または記録する関数です。</span><span class="sxs-lookup"><span data-stu-id="4ca29-166">Note that `reportError` is a function that shows or logs an error.</span></span>

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

## <a name="error-handling"></a><span data-ttu-id="4ca29-167">エラー処理</span><span class="sxs-lookup"><span data-stu-id="4ca29-167">Error handling</span></span>

<span data-ttu-id="4ca29-168">一部のシナリオでは、Office はリボンを更新できず、エラーを返します。</span><span class="sxs-lookup"><span data-stu-id="4ca29-168">In some scenarios, Office is unable to update the ribbon and will return an error.</span></span> <span data-ttu-id="4ca29-169">たとえば、アドインがアップグレードされ、アップグレードされたアドインに異なるカスタム アドイン コマンドのセットがある場合は、Office アプリケーションを閉じてから、もう一度開く必要があります。</span><span class="sxs-lookup"><span data-stu-id="4ca29-169">For example, if the add-in is upgraded and the upgraded add-in has a different set of custom add-in commands, then the Office application must be closed and reopened.</span></span> <span data-ttu-id="4ca29-170">それまでの間、`requestUpdate` メソッドは `HostRestartNeeded` エラーを返します。</span><span class="sxs-lookup"><span data-stu-id="4ca29-170">Until it is, the `requestUpdate` method will return the error `HostRestartNeeded`.</span></span> <span data-ttu-id="4ca29-171">このエラーの処理方法の例を次に示します。</span><span class="sxs-lookup"><span data-stu-id="4ca29-171">The following is an example of how to handle this error.</span></span> <span data-ttu-id="4ca29-172">この場合、`reportError` メソッドがユーザーにエラーを表示します。</span><span class="sxs-lookup"><span data-stu-id="4ca29-172">In this case, the `reportError` method displays the error to the user.</span></span>

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

## <a name="test-for-platform-support-with-requirement-sets"></a><span data-ttu-id="4ca29-173">要件セットを使用したプラットフォーム サポートのテスト</span><span class="sxs-lookup"><span data-stu-id="4ca29-173">Test for platform support with requirement sets</span></span>

<span data-ttu-id="4ca29-p122">要件セットは、API メンバーの名前付きグループです。Office アドインは、マニフェストで指定されている要件セットを使用するか、ランタイム チェックを使用して、Office ホストがアドインに必要な API をサポートしているかどうかを判別します。詳しくは、「[Office のバージョンと要件セット](../develop/office-versions-and-requirement-sets.md)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="4ca29-p122">Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](../develop/office-versions-and-requirement-sets.md).</span></span>

<span data-ttu-id="4ca29-177">API を有効化/無効化するには、次の要件セットをサポートしている必要があります。</span><span class="sxs-lookup"><span data-stu-id="4ca29-177">The enable/disable APIs require support of the following requirement set:</span></span>

- [<span data-ttu-id="4ca29-178">AddinCommands 1.1</span><span class="sxs-lookup"><span data-stu-id="4ca29-178">AddinCommands 1.1</span></span>](../reference/requirement-sets/add-in-commands-requirement-sets.md)
