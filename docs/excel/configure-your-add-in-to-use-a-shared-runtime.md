---
ms.date: 02/20/2020
title: ブラウザーのランタイムを共有するように Excel アドインを構成する (プレビュー)
ms.prod: excel
description: Excel アドインを構成して、ブラウザーのランタイムを共有し、同じランタイムでリボン、作業ウィンドウ、カスタム関数のコードを実行できるようにします。
localization_priority: Priority
ms.openlocfilehash: 7fad97039d5a463978ca9a2d450af9bb698b49b5
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/17/2020
ms.locfileid: "42719330"
---
# <a name="configure-your-excel-add-in-to-use-a-shared-javascript-runtime-preview"></a><span data-ttu-id="de22f-103">共有の JavaScript ランタイムを使用するように Excel アドインを構成する (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="de22f-103">Configure your Excel add-in to use a shared JavaScript runtime (preview)</span></span>

[!include[Running custom functions in a shared runtime note](../includes/excel-shared-runtime-preview-note.md)]

<span data-ttu-id="de22f-104">Windows または Mac で Excel を実行する場合、アドインは、リボン ボタン、カスタム関数、作業ウィンドウのコードを別の JavaScript ランタイム環境で実行します。</span><span class="sxs-lookup"><span data-stu-id="de22f-104">When running Excel on Windows or Mac, your add-in will run code for ribbon buttons, custom functions, and the task pane in separate JavaScript runtime environments.</span></span> <span data-ttu-id="de22f-105">これにより、グローバル データを簡単に共有できない、カスタム関数からすべての CORS 機能にアクセスできないなどの制限が発生します。</span><span class="sxs-lookup"><span data-stu-id="de22f-105">This creates limitations such as not being able to easily share global data, and not being able to access all CORS functionality from a custom function.</span></span>

<span data-ttu-id="de22f-106">ただし、Excel アドインを構成すれば、共有の JavaScript ランタイムでコードを共有できるようになります。</span><span class="sxs-lookup"><span data-stu-id="de22f-106">However, you can configure your Excel add-in to share code in a shared JavaScript runtime.</span></span> <span data-ttu-id="de22f-107">これにより、アドイン間での調整が容易になり、アドインのすべての部分から DOM や CORS にアクセスできます。</span><span class="sxs-lookup"><span data-stu-id="de22f-107">This enables better coordination across your add-in and access to the DOM and CORS from all parts of your add-in.</span></span> <span data-ttu-id="de22f-108">また、ドキュメントを開いているときにコードを実行したり、作業ウィンドウが閉じた状態でコードを実行したりできます。</span><span class="sxs-lookup"><span data-stu-id="de22f-108">It also enables you to run code when the document opens, or to run code while the task pane is closed.</span></span> <span data-ttu-id="de22f-109">共有ランタイムが使用できるようにアドインを構成するには、この記事の手順に従います。</span><span class="sxs-lookup"><span data-stu-id="de22f-109">To configure your add-in to use a shared runtime, follow the instructions in this article.</span></span>

## <a name="create-the-add-in-project"></a><span data-ttu-id="de22f-110">アドイン プロジェクトの作成</span><span class="sxs-lookup"><span data-stu-id="de22f-110">Create the add-in project</span></span>

<span data-ttu-id="de22f-111">新しいプロジェクトを開始する場合は、次の手順に従って、Yeoman ジェネレーターを使って Excel アドインを作成します。</span><span class="sxs-lookup"><span data-stu-id="de22f-111">If you are starting a new project, follow these steps to use the Yeoman generator to create an Excel add-in project.</span></span> <span data-ttu-id="de22f-112">次のコマンドを実行し、プロンプトに次の回答を入力します。</span><span class="sxs-lookup"><span data-stu-id="de22f-112">Run the following command and then answer the prompts with the following answers:</span></span>

```command line
yo office
```

- <span data-ttu-id="de22f-113">プロジェクトの種類を選択する:  **Excel カスタム関数アドイン プロジェクト**</span><span class="sxs-lookup"><span data-stu-id="de22f-113">Choose a project type: **Excel Custom Functions Add-in project**</span></span>
- <span data-ttu-id="de22f-114">スクリプトの種類を選択する:  **JavaScript**</span><span class="sxs-lookup"><span data-stu-id="de22f-114">Choose a script type: **JavaScript**</span></span>
- <span data-ttu-id="de22f-115">アドインの名前を何にしますか?  **個人用 Office アドイン**</span><span class="sxs-lookup"><span data-stu-id="de22f-115">What do you want to name your add-in? **My Office Add-in**</span></span>

![アドイン プロジェクトを作成するための Office からのプロンプトへ応答するスクリーンショット。](../images/yo-office-excel-project.png)

<span data-ttu-id="de22f-117">ウィザードを完了すると、ジェネレーターによってプロジェクトが作成され、サポートしているノード コンポーネントがインストールされます。</span><span class="sxs-lookup"><span data-stu-id="de22f-117">After you complete the wizard, the generator creates the project and installs supporting Node components.</span></span>

## <a name="configure-the-manifest"></a><span data-ttu-id="de22f-118">マニフェストを構成する</span><span class="sxs-lookup"><span data-stu-id="de22f-118">Configure the manifest</span></span>

<span data-ttu-id="de22f-119">新規または既存のプロジェクトで共有ランタイムが使用できるように構成するには、次の手順を実行します。</span><span class="sxs-lookup"><span data-stu-id="de22f-119">Follow these steps for a new or existing project to configure it to use a shared runtime.</span></span>

1. <span data-ttu-id="de22f-120">Visual Studio Code を開始して [**個人用 Office アドイン**] プロジェクトを開きます。</span><span class="sxs-lookup"><span data-stu-id="de22f-120">Start Visual Studio Code and open the **My Office Add-in** project.</span></span>
2. <span data-ttu-id="de22f-121">
            **manifest.xml** ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="de22f-121">Open the **manifest.xml** file.</span></span>
3. <span data-ttu-id="de22f-122">`<VersionOverrides>` セクションを探し、次の `<Runtimes>` セクションを追加します。</span><span class="sxs-lookup"><span data-stu-id="de22f-122">Find the `<VersionOverrides>` section, and add the following `<Runtimes>` section.</span></span> <span data-ttu-id="de22f-123">作業ウィンドウを閉じてもカスタム関数が引き続き機能するように、有効期間は**長く**する必要があります。</span><span class="sxs-lookup"><span data-stu-id="de22f-123">The lifetime needs to be **long** so that the custom functions can still work even when the task pane is closed.</span></span> <span data-ttu-id="de22f-124">resid は `ContosoAddin.Url` で、後述のリソースのセクションの文字列を参照します。</span><span class="sxs-lookup"><span data-stu-id="de22f-124">The resid is `ContosoAddin.Url` which references a string in the resources section later.</span></span> <span data-ttu-id="de22f-125">resid には任意の値を使用できますが、アドイン要素のその他の要素の resid と一致している必要があります。</span><span class="sxs-lookup"><span data-stu-id="de22f-125">You can use any resid value you want, but it should match the resid of the other elements in your add-in elements.</span></span>

   ```xml
   <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
     <Hosts>
       <Host xsi:type="Workbook">
       <Runtimes>
         <Runtime resid="ContosoAddin.Url" lifetime="long" />
       </Runtimes>
       <AllFormFactors>
   ```

4. <span data-ttu-id="de22f-126">`<Page>` 要素で、ソースの場所を **Functions.Page.Url** から **ContosoAddin.Url** に変更します。</span><span class="sxs-lookup"><span data-stu-id="de22f-126">In the `<Page>` element, change the source location from **Functions.Page.Url** to **ContosoAddin.Url**.</span></span> <span data-ttu-id="de22f-127">この resid は、`<Runtime>` resid の要素と一致しています。</span><span class="sxs-lookup"><span data-stu-id="de22f-127">This resid matches the `<Runtime>` resid element.</span></span> <span data-ttu-id="de22f-128">カスタム関数がない場合は、**Page** エントリがないため、この手順は省略できます。</span><span class="sxs-lookup"><span data-stu-id="de22f-128">Note that if you don't have custom functions, you will not have a **Page** entry and can skip this step.</span></span>

   ```xml
   <AllFormFactors>
   ...
   <Page>
   <SourceLocation resid="ContosoAddin.Url"/>
   </Page>
   ...
   ```

5. <span data-ttu-id="de22f-129">`<DesktopFormFactor>` セクションで、**FunctionFile** を **Commands.Url** から **ContosoAddin.Url** を使用するように変更します。</span><span class="sxs-lookup"><span data-stu-id="de22f-129">In the `<DesktopFormFactor>` section, change the **FunctionFile** from **Commands.Url** to use **ContosoAddin.Url**.</span></span> <span data-ttu-id="de22f-130">アクション コマンドがない場合は、**FunctionFile** エントリがないため、この手順は省略できます。</span><span class="sxs-lookup"><span data-stu-id="de22f-130">Note that if you don't have action commands, you won't have a **FunctionFile** entry, and can skip this step.</span></span>

   ```xml
   <DesktopFormFactor>
   <GetStarted>
   ...
   </GetStarted>
   <FunctionFile resid="ContosoAddin.Url"/>
   ```

6. <span data-ttu-id="de22f-131">`<Action>` セクションで、ソースの場所を **Taskpane.Url** から **ContosoAddin.Url** に変更します。</span><span class="sxs-lookup"><span data-stu-id="de22f-131">In the `<Action>` section, change the source location from **Taskpane.Url** to **ContosoAddin.Url**.</span></span> <span data-ttu-id="de22f-132">作業ウィンドウがない場合は、**ShowTaskpane** アクションがないため、この手順は省略できます。</span><span class="sxs-lookup"><span data-stu-id="de22f-132">Note that if you don't have a task pane, you won't have a **ShowTaskpane** action, and can skip this step.</span></span>

   ```xml
   <Action xsi:type="ShowTaskpane">
   <TaskpaneId>ButtonId1</TaskpaneId>
   <SourceLocation resid="ContosoAddin.Url"/>
   </Action>
   ```

7. <span data-ttu-id="de22f-133">**taskpane.html** を指す **ContosoAddin.Url** の新しい **Url ID** を追加します。</span><span class="sxs-lookup"><span data-stu-id="de22f-133">Add a new **Url id** for **ContosoAddin.Url** that points to **taskpane.html**.</span></span>

   ```xml
   <bt:Urls>
   <bt:Url id="Functions.Script.Url" DefaultValue="https://localhost:3000/dist/functions.js"/>
   ...
   <bt:Url id="ContosoAddin.Url" DefaultValue="https://localhost:3000/taskpane.html"/>
   ...
   ```

8. <span data-ttu-id="de22f-134">変更を保存してプロジェクトを再ビルドします。</span><span class="sxs-lookup"><span data-stu-id="de22f-134">Save your changes and rebuild the project.</span></span>

   ```command line
   npm run build
   ```

## <a name="runtime-lifetime"></a><span data-ttu-id="de22f-135">ランタイムの有効期間</span><span class="sxs-lookup"><span data-stu-id="de22f-135">Runtime lifetime</span></span>

<span data-ttu-id="de22f-136">`Runtime` 要素を追加するときに、有効期間も `long` または `short` の値で指定します。</span><span class="sxs-lookup"><span data-stu-id="de22f-136">When you add the `Runtime` element, you also specify a lifetime with a value of `long` or `short`.</span></span> <span data-ttu-id="de22f-137">この値を `long` に設定すると、ドキュメントを開くとアドインを起動したり、作業ウィンドウを閉じた後にコードを継続して実行したり、カスタム関数から CORS および DOM を使用したりできます。</span><span class="sxs-lookup"><span data-stu-id="de22f-137">Set this value to `long` to take advantage of features such as starting your add-in when the document opens, continuing to run code after the task pane is closed, or using CORS and DOM from custom functions.</span></span>

<span data-ttu-id="de22f-138">この値を `short`に設定すると、アドインは既定の動作と同様に動作します。</span><span class="sxs-lookup"><span data-stu-id="de22f-138">If you set this value to `short`, the add-in will behave similar to the default behavior.</span></span> <span data-ttu-id="de22f-139">いずれかのリボン ボタンを押すとアドインが開始しますが、リボン ハンドラーの実行を完了した後で、アドインが終了する場合があります。</span><span class="sxs-lookup"><span data-stu-id="de22f-139">The add-in will start when one of your ribbon buttons is pressed, but it may shut down after your ribbon handler is done running.</span></span> <span data-ttu-id="de22f-140">同様に、作業ウィンドウを開くとアドインが起動します。ただし、作業ウィンドウを閉じると、アドインが終了する場合があります。</span><span class="sxs-lookup"><span data-stu-id="de22f-140">Similarly your add-in will start when the task pane is opened, but it may shut down when the task pane is closed.</span></span>

```xml
<Runtimes>
  <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="multiple-task-panes"></a><span data-ttu-id="de22f-141">複数の作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="de22f-141">Multiple task panes</span></span>

<span data-ttu-id="de22f-142">共有ランタイムを使用する予定がある場合は、複数の作業ウィンドウを使用するようにアドインを設計しないでください。</span><span class="sxs-lookup"><span data-stu-id="de22f-142">Don't design your add-in to use multiple task panes if you are planning to use shared runtime.</span></span> <span data-ttu-id="de22f-143">共有ランタイムは、1 つの作業ウィンドウのみサポートします。</span><span class="sxs-lookup"><span data-stu-id="de22f-143">Shared runtime only supports the use of one task pane.</span></span> <span data-ttu-id="de22f-144">`<TaskpaneID>` のない作業ウィンドウは、別の作業ウィンドウとして扱われますのでご注意ください。</span><span class="sxs-lookup"><span data-stu-id="de22f-144">Note that any task pane without a `<TaskpaneID>` is considered a different task pane.</span></span>

## <a name="next-steps"></a><span data-ttu-id="de22f-145">次の手順</span><span class="sxs-lookup"><span data-stu-id="de22f-145">Next steps</span></span>

<span data-ttu-id="de22f-146">次の記事を参照して、共有ランタイムの機能をいくつかお試しください。</span><span class="sxs-lookup"><span data-stu-id="de22f-146">Now try out some features of the shared runtime by taking a look at the following articles.</span></span>

- [<span data-ttu-id="de22f-147">カスタム関数から Excel API を呼び出す</span><span class="sxs-lookup"><span data-stu-id="de22f-147">Call Excel APIs from a custom function</span></span>](call-excel-apis-from-custom-function.md)

## <a name="see-also"></a><span data-ttu-id="de22f-148">関連項目</span><span class="sxs-lookup"><span data-stu-id="de22f-148">See also</span></span>

- [<span data-ttu-id="de22f-149">概要: 共有の JavaScript ランタイムでアドイン コードを実行する (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="de22f-149">Overview: Run your add-in code in a shared JavaScript runtime (preview)</span></span>](custom-functions-shared-overview.md)
