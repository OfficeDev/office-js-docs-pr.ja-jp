---
ms.date: 06/14/2021
title: 共有 JavaScript ランタイムを使用するように Office アドインを構成する
ms.prod: non-product-specific
description: 共有 JavaScript ランタイムを使用して、追加のリボン、作業ウィンドウ、およびカスタム関数機能をサポートするように Office アドインを構成します。
localization_priority: Priority
ms.openlocfilehash: ecde9a5564761b2dd902596f09db156332b5af4f
ms.sourcegitcommit: 4fa952f78be30d339ceda3bd957deb07056ca806
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/16/2021
ms.locfileid: "52961259"
---
# <a name="configure-your-office-add-in-to-use-a-shared-javascript-runtime"></a><span data-ttu-id="05232-103">共有 JavaScript ランタイムを使用するように Office アドインを構成する</span><span class="sxs-lookup"><span data-stu-id="05232-103">Configure your Office Add-in to use a shared JavaScript runtime</span></span>

[!include[Shared JavaScript runtime requirements](../includes/shared-runtime-requirements-note.md)]

<span data-ttu-id="05232-104">単一の共有 JavaScript ランタイム (共有ランタイムとも呼ばれる) ですべてのコードを実行するように Office アドインを構成できます。</span><span class="sxs-lookup"><span data-stu-id="05232-104">You can configure your Office Add-in to run all of its code in a single shared JavaScript runtime (also known as a shared runtime).</span></span> <span data-ttu-id="05232-105">これにより、アドイン間での調整が容易になり、アドインのすべての部分から DOM や CORS にアクセスできます。</span><span class="sxs-lookup"><span data-stu-id="05232-105">This enables better coordination across your add-in and access to the DOM and CORS from all parts of your add-in.</span></span> <span data-ttu-id="05232-106">また、ドキュメントを開いたときにコードを実行したり、リボン ボタンを有効または無効にするなどの追加機能も有効にできます。</span><span class="sxs-lookup"><span data-stu-id="05232-106">It also enables additional features such as running code when the document opens, or enabling or disabling ribbon buttons.</span></span> <span data-ttu-id="05232-107">共有 JavaScript ランタイムが使用できるようにアドインを構成するには、この記事の手順に従います。</span><span class="sxs-lookup"><span data-stu-id="05232-107">To configure your add-in to use a shared JavaScript runtime, follow the instructions in this article.</span></span>

## <a name="create-the-add-in-project"></a><span data-ttu-id="05232-108">アドイン プロジェクトの作成</span><span class="sxs-lookup"><span data-stu-id="05232-108">Create the add-in project</span></span>

<span data-ttu-id="05232-109">新しいプロジェクトを開始する場合は、次の手順に従って、[Office アドイン用の Yeoman ジェネレーター](https://github.com/OfficeDev/generator-office)を使って Excel アドインまたは PowerPoint アドインを作成します。</span><span class="sxs-lookup"><span data-stu-id="05232-109">If you are starting a new project, follow these steps to use the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office) to create an Excel or PowerPoint add-in project.</span></span>

<span data-ttu-id="05232-110">次のいずれかの操作を行います。</span><span class="sxs-lookup"><span data-stu-id="05232-110">Do one of the following:</span></span>

- <span data-ttu-id="05232-111">カスタム関数を使用して Excel アドインを生成するには、コマンド `yo office --projectType excel-functions --name 'Excel shared runtime add-in' --host excel --js true` を実行します。</span><span class="sxs-lookup"><span data-stu-id="05232-111">To generate an Excel add-in with custom functions, run the command `yo office --projectType excel-functions --name 'Excel shared runtime add-in' --host excel --js true`.</span></span>

    <span data-ttu-id="05232-112">または</span><span class="sxs-lookup"><span data-stu-id="05232-112">or</span></span>

- <span data-ttu-id="05232-113">PowerPoint アドインを生成するには、コマンド `yo office --projectType taskpane --name 'PowerPoint shared runtime add-in' --host powerpoint --js true` を実行します。</span><span class="sxs-lookup"><span data-stu-id="05232-113">To generate a PowerPoint add-in, run the command `yo office --projectType taskpane --name 'PowerPoint shared runtime add-in' --host powerpoint --js true`.</span></span>

<span data-ttu-id="05232-114">ジェネレーターはプロジェクトを作成し、サポートしているノード コンポーネントをインストールします。</span><span class="sxs-lookup"><span data-stu-id="05232-114">The generator will create the project and install supporting Node components.</span></span>

> [!NOTE]
> <span data-ttu-id="05232-115">この記事の手順を使用して、共有ランタイムを使用するように既存の Visual Studio プロジェクトを更新することもできます。</span><span class="sxs-lookup"><span data-stu-id="05232-115">You can also use the steps in this article to update an existing Visual Studio project to use the shared runtime.</span></span> <span data-ttu-id="05232-116">ただし、マニフェストの XML スキーマの更新が必要になる場合があります。</span><span class="sxs-lookup"><span data-stu-id="05232-116">However, you may need to update the XML schemas for the manifest.</span></span> <span data-ttu-id="05232-117">詳細については、「[Office アドインでの開発エラーのトラブルシューティング](../testing/troubleshoot-development-errors.md#manifest-schema-validation-errors-in-visual-studio-projects)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="05232-117">For more information, see [Troubleshoot development errors with Office Add-ins](../testing/troubleshoot-development-errors.md#manifest-schema-validation-errors-in-visual-studio-projects).</span></span>

## <a name="configure-the-manifest"></a><span data-ttu-id="05232-118">マニフェストを構成する</span><span class="sxs-lookup"><span data-stu-id="05232-118">Configure the manifest</span></span>

<span data-ttu-id="05232-119">新規または既存のプロジェクトで共有ランタイムが使用できるように構成するには、次の手順を実行します。</span><span class="sxs-lookup"><span data-stu-id="05232-119">Follow these steps for a new or existing project to configure it to use a shared runtime.</span></span> <span data-ttu-id="05232-120">これらの手順は、[Office アドイン用の Yeoman ジェネレーター](https://github.com/OfficeDev/generator-office)を使用してプロジェクトを生成したことを前提としています。</span><span class="sxs-lookup"><span data-stu-id="05232-120">These steps assume you have generated your project using the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office).</span></span>

1. <span data-ttu-id="05232-121">Visual Studio Code を起動し、生成した Excel または PowerPoint アドイン プロジェクトを開きます。</span><span class="sxs-lookup"><span data-stu-id="05232-121">Start Visual Studio Code and open the Excel or PowerPoint add-in project you generated.</span></span>
1. <span data-ttu-id="05232-122">**manifest.xml** ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="05232-122">Open the **manifest.xml** file.</span></span>
1. <span data-ttu-id="05232-123">Excel アドインを生成した場合は、要件セクションを更新して、カスタム関数ランタイムの代わりに[共有ランタイム](../reference/requirement-sets/shared-runtime-requirement-sets.md)を使用します。</span><span class="sxs-lookup"><span data-stu-id="05232-123">If you generated an Excel add-in, update the requirements section to use the [shared runtime](../reference/requirement-sets/shared-runtime-requirement-sets.md) instead of the custom function runtime.</span></span> <span data-ttu-id="05232-124">XML は、次のようになります。</span><span class="sxs-lookup"><span data-stu-id="05232-124">The XML should appear as follows.</span></span>

    ```xml
    <Hosts>
      <Host Name="Workbook"/>
    </Hosts>
    <Requirements>
      <Sets DefaultMinVersion="1.1">
        <Set Name="SharedRuntime" MinVersion="1.1"/>
      </Sets>
    </Requirements>
    <DefaultSettings>
    ```

1. <span data-ttu-id="05232-125">`<VersionOverrides>` セクションを見つけて、`<Host ...>` タグのすぐ内側に次の `<Runtimes>` セクションを追加します。</span><span class="sxs-lookup"><span data-stu-id="05232-125">Find the `<VersionOverrides>` section and add the following `<Runtimes>` section just inside the `<Host ...>` tag.</span></span> <span data-ttu-id="05232-126">作業ウィンドウを閉じてもアドイン コードを実行できるように、有効期間は **長く** する必要があります。</span><span class="sxs-lookup"><span data-stu-id="05232-126">The lifetime needs to be **long** so that your add-in code can run even when the task pane is closed.</span></span> <span data-ttu-id="05232-127">`resid` 値は **Taskpane.Url** で、**manifest.xml** ファイルの下部付近の ` <bt:Urls>` セクションで指定された **taskpane.html** ファイルの場所を参照します。</span><span class="sxs-lookup"><span data-stu-id="05232-127">The `resid` value is **Taskpane.Url**, which references the **taskpane.html** file location specified in the ` <bt:Urls>` section near the bottom of the **manifest.xml** file.</span></span>

   ```xml
   <VersionOverrides ...>
     <Hosts>
       <Host ...>
       <Runtimes>
         <Runtime resid="Taskpane.Url" lifetime="long" />
       </Runtimes>
       ...
   ```

1. <span data-ttu-id="05232-128">カスタム関数を使用して Excel アドインを生成した場合は、`<Page>` 要素を見つけます。</span><span class="sxs-lookup"><span data-stu-id="05232-128">If you generated an Excel add-in with custom functions, find the `<Page>` element.</span></span> <span data-ttu-id="05232-129">次に、ソースの場所を **Functions.Page.Url** から **Taskpane.Url** に変更します。</span><span class="sxs-lookup"><span data-stu-id="05232-129">Then change the source location from **Functions.Page.Url** to **Taskpane.Url**.</span></span>

   ```xml
   <AllFormFactors>
   ...
   <Page>
     <SourceLocation resid="Taskpane.Url"/>
   </Page>
   ...
   ```

1. <span data-ttu-id="05232-130">`<FunctionFile ...>` タグを見つけて、`resid` を **Commands.Url** から **Taskpane.Url** に変更します。</span><span class="sxs-lookup"><span data-stu-id="05232-130">Find the `<FunctionFile ...>` tag and change the `resid` from **Commands.Url** to  **Taskpane.Url**.</span></span> <span data-ttu-id="05232-131">アクション コマンドがない場合は、**FunctionFile** エントリがないため、この手順は省略できます。</span><span class="sxs-lookup"><span data-stu-id="05232-131">Note that if you don't have action commands, you won't have a **FunctionFile** entry, and can skip this step.</span></span>

    ```xml
    </GetStarted>
    ...
    <FunctionFile resid="Taskpane.Url"/>
    ...
    ```

1. <span data-ttu-id="05232-132">**manifest.xml** ファイルを保存します。</span><span class="sxs-lookup"><span data-stu-id="05232-132">Save the **manifest.xml** file.</span></span>

## <a name="configure-the-webpackconfigjs-file"></a><span data-ttu-id="05232-133">webpack.config.js ファイルを構成する</span><span class="sxs-lookup"><span data-stu-id="05232-133">Configure the webpack.config.js file</span></span>

<span data-ttu-id="05232-134">**webpack.config.js** は、複数のランタイム ローダーをビルドします。</span><span class="sxs-lookup"><span data-stu-id="05232-134">The **webpack.config.js** will build multiple runtime loaders.</span></span> <span data-ttu-id="05232-135">**taskpane.html** ファイルを介して共有 JavaScript ランタイムのみを読み込むように変更する必要があります。</span><span class="sxs-lookup"><span data-stu-id="05232-135">You need to modify it to load only the shared JavaScript runtime via the **taskpane.html** file.</span></span>

1. <span data-ttu-id="05232-136">Visual Studio Code を起動し、生成した Excel または PowerPoint アドイン プロジェクトを開きます。</span><span class="sxs-lookup"><span data-stu-id="05232-136">Start Visual Studio Code and open the Excel or PowerPoint add-in project you generated.</span></span>
1. <span data-ttu-id="05232-137">**webpack.config.js** ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="05232-137">Open the **webpack.config.js** file.</span></span>
1. <span data-ttu-id="05232-138">**webpack.config.js** ファイルに次の **functions.html** プラグイン コードが含まれている場合は、それを削除します。</span><span class="sxs-lookup"><span data-stu-id="05232-138">If your **webpack.config.js** file has the following **functions.html** plugin code, remove it.</span></span>

    ```javascript
    new HtmlWebpackPlugin({
        filename: "functions.html",
        template: "./src/functions/functions.html",
        chunks: ["polyfill", "functions"]
      })
    ```

1. <span data-ttu-id="05232-139">**webpack.config.js** ファイルに次の **commands.html** プラグイン コードが含まれている場合は、それを削除します。</span><span class="sxs-lookup"><span data-stu-id="05232-139">If your **webpack.config.js** file has the following **commands.html** plugin code, remove it.</span></span>

    ```javascript
    new HtmlWebpackPlugin({
        filename: "commands.html",
        template: "./src/commands/commands.html",
        chunks: ["polyfill", "commands"]
      })
    ```

1. <span data-ttu-id="05232-140">プロジェクトで **関数** チャンクまたは **コマンド** チャンクのいずれかを使用した場合は、次に示すようにそれらをチャンク リストに追加します (次のコードは、プロジェクトで両方のチャンクを使用した場合のものです)。</span><span class="sxs-lookup"><span data-stu-id="05232-140">If your project used either the **functions** or **commands** chunks, add them to the chunks list as shown next (the following code is for if your project used both chunks).</span></span>

    ```javascript
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["polyfill", "taskpane", "commands", "functions"]
      })
    ```

1. <span data-ttu-id="05232-141">変更を保存してプロジェクトを再ビルドします。</span><span class="sxs-lookup"><span data-stu-id="05232-141">Save your changes and rebuild the project.</span></span>

   ```command line
   npm run build
   ```

> [!NOTE]
> <span data-ttu-id="05232-142">プロジェクトに **functions.html** ファイルまたは **commands.html** ファイルがある場合は、それらを削除できます。</span><span class="sxs-lookup"><span data-stu-id="05232-142">If your project has a **functions.html** file or **commands.html** file, they can be removed.</span></span> <span data-ttu-id="05232-143">**taskpane.html** は、先ほど行った webpack の更新を介して、**functions.js** および **commands.js** コードを共有 JavaScript ランタイムに読み込みます。</span><span class="sxs-lookup"><span data-stu-id="05232-143">The **taskpane.html** will load the **functions.js** and **commands.js** code into the shared JavaScript runtime via the webpack updates you just made.</span></span>

## <a name="test-your-office-add-in-changes"></a><span data-ttu-id="05232-144">Office アドインの変更をテストする</span><span class="sxs-lookup"><span data-stu-id="05232-144">Test your Office Add-in changes</span></span>

<span data-ttu-id="05232-145">共有 JavaScript ランタイムが正しく使用されていることを確認するには、次の手順を実行します。</span><span class="sxs-lookup"><span data-stu-id="05232-145">You can confirm that you are using the shared JavaScript runtime correctly by using the following instructions.</span></span>

1. <span data-ttu-id="05232-146">**manifest.xml** ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="05232-146">Open the **manifest.xml** file.</span></span>
1. <span data-ttu-id="05232-147">`<Control xsi:type="Button" id="TaskpaneButton">` セクションを探し、次の `<Action ...>` XML を変更します。</span><span class="sxs-lookup"><span data-stu-id="05232-147">Find the `<Control xsi:type="Button" id="TaskpaneButton">` section and change the following `<Action ...>` XML.</span></span>

    <span data-ttu-id="05232-148">送信元:</span><span class="sxs-lookup"><span data-stu-id="05232-148">from:</span></span>

    ```xml
    <Action xsi:type="ShowTaskpane">
      <TaskpaneId>ButtonId1</TaskpaneId>
      <SourceLocation resid="Taskpane.Url"/>
    </Action>
    ```

    <span data-ttu-id="05232-149">変更後:</span><span class="sxs-lookup"><span data-stu-id="05232-149">to:</span></span>

    ```xml
    <Action xsi:type="ExecuteFunction">
      <FunctionName>action</FunctionName>
    </Action>
    ```

1. <span data-ttu-id="05232-150">**./src/commands/commands.js** ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="05232-150">Open the **./src/commands/commands.js** file.</span></span>
1. <span data-ttu-id="05232-151">**アクション** 関数を以下のコードに置き換えます。</span><span class="sxs-lookup"><span data-stu-id="05232-151">Replace the **action** function with the code below.</span></span> <span data-ttu-id="05232-152">これにより、関数が更新され、作業ウィンドウ ボタンが開いて変更され、カウンターが増加されます。</span><span class="sxs-lookup"><span data-stu-id="05232-152">This will update the function to open and modify the task pane button to increment a counter.</span></span> <span data-ttu-id="05232-153">コマンドから作業ウィンドウ DOM を開いてアクセスすることは、共有 JavaScript ランタイムでのみ機能します。</span><span class="sxs-lookup"><span data-stu-id="05232-153">Opening and accessing the task pane DOM from a command only works with the shared JavaScript runtime.</span></span>

    ```javascript
    var _count=0;
    
    function action(event) {
      // Your code goes here.
      _count++;
      Office.addin.showAsTaskpane();
      document.getElementById("run").textContent="Go"+_count;
    
      // Be sure to indicate when the add-in command function is complete.
      event.completed();
    }
    ```

1. <span data-ttu-id="05232-154">変更を保存してプロジェクトを実行します。</span><span class="sxs-lookup"><span data-stu-id="05232-154">Save your changes and run the project.</span></span>

   ```command line
   npm start
   ```

<span data-ttu-id="05232-155">アドイン ボタンを選択するたびに、[**実行**] ボタンのテキストが [**移動**] に変更され、その後にカウンターが増加されます。</span><span class="sxs-lookup"><span data-stu-id="05232-155">Each time you select the add-ins button, it will change the **run** button text to **go** and increment a counter after it.</span></span>

## <a name="runtime-lifetime"></a><span data-ttu-id="05232-156">ランタイムの有効期間</span><span class="sxs-lookup"><span data-stu-id="05232-156">Runtime lifetime</span></span>

<span data-ttu-id="05232-157">`Runtime` 要素を追加するときに、有効期間も `long` または `short` の値で指定します。</span><span class="sxs-lookup"><span data-stu-id="05232-157">When you add the `Runtime` element, you also specify a lifetime with a value of `long` or `short`.</span></span> <span data-ttu-id="05232-158">この値を `long` に設定すると、ドキュメントを開くとアドインを起動したり、作業ウィンドウを閉じた後にコードを継続して実行したり、カスタム関数から CORS および DOM を使用したりできます。</span><span class="sxs-lookup"><span data-stu-id="05232-158">Set this value to `long` to take advantage of features such as starting your add-in when the document opens, continuing to run code after the task pane is closed, or using CORS and DOM from custom functions.</span></span>

> [!NOTE]
> <span data-ttu-id="05232-159">既定の有効期間の値は `short` ですが、Excel アドインでは `long` を使うことをお勧めします。この例でランタイムを `short` に設定した場合、いずれかのリボン ボタンを押したときに Excel アドインが起動しますが、リボン ハンドラーの実行が完了するとアドインが終了することがあります。</span><span class="sxs-lookup"><span data-stu-id="05232-159">The default lifetime value is `short`, but we recommend using `long` in Excel add-ins. If you set your runtime to `short` in this example, your Excel add-in will start when one of your ribbon buttons is pressed, but it may shut down after your ribbon handler is done running.</span></span> <span data-ttu-id="05232-160">同様に、作業ウィンドウを開くとアドインが起動します。ただし、作業ウィンドウを閉じると、アドインが終了する場合があります。</span><span class="sxs-lookup"><span data-stu-id="05232-160">Similarly, your add-in will start when the task pane is opened, but it may shut down when the task pane is closed.</span></span>

```xml
<Runtimes>
  <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

> [!NOTE]
> <span data-ttu-id="05232-161">アドインにマニフェストの `Runtimes` 要素が含まれている場合 (共有ランタイムに必要)、Windows または Microsoft 365 のバージョンに関係なく、Internet Explorer 11 が使用されます。</span><span class="sxs-lookup"><span data-stu-id="05232-161">If your add-in includes the `Runtimes` element in the manifest (required for a shared runtime), it uses Internet Explorer 11 regardless of the Windows or Microsoft 365 version.</span></span> <span data-ttu-id="05232-162">詳細については、「[ランタイム](../reference/manifest/runtimes.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="05232-162">For more information, see [Runtimes](../reference/manifest/runtimes.md).</span></span>

## <a name="about-the-shared-javascript-runtime"></a><span data-ttu-id="05232-163">共有 JavaScript ランタイムについて</span><span class="sxs-lookup"><span data-stu-id="05232-163">About the shared JavaScript runtime</span></span>

<span data-ttu-id="05232-164">Windows または Mac で、アドインは、リボン ボタン、カスタム関数、作業ウィンドウのコードを別の JavaScript ランタイム環境で実行します。</span><span class="sxs-lookup"><span data-stu-id="05232-164">On Windows or Mac, your add-in will run code for ribbon buttons, custom functions, and the task pane in separate JavaScript runtime environments.</span></span> <span data-ttu-id="05232-165">これにより、グローバル データを簡単に共有できない、カスタム関数からすべての CORS 機能にアクセスできないなどの制限が発生します。</span><span class="sxs-lookup"><span data-stu-id="05232-165">This creates limitations such as not being able to easily share global data, and not being able to access all CORS functionality from a custom function.</span></span>

<span data-ttu-id="05232-166">ただし、Office アドインを構成すれば、同じ JavaScript ランタイム (共有ランタイムとも呼ばれる) でコードを共有できるようになります。</span><span class="sxs-lookup"><span data-stu-id="05232-166">However, you can configure your Office Add-in to share code in the same JavaScript runtime (also referred to as a shared runtime).</span></span> <span data-ttu-id="05232-167">これにより、アドイン間での調整が容易になり、アドインのすべての部分から、作業ウィンドウの DOM や CORS にアクセスできます。</span><span class="sxs-lookup"><span data-stu-id="05232-167">This enables better coordination across your add-in and access to the task pane DOM and CORS from all parts of your add-in.</span></span>

<span data-ttu-id="05232-168">共有ランタイムを構成すると、次のシナリオが可能になります。</span><span class="sxs-lookup"><span data-stu-id="05232-168">Configuring a shared runtime enables the following scenarios.</span></span>

- <span data-ttu-id="05232-169">Office アドインは、追加の UI 機能を使用できます。</span><span class="sxs-lookup"><span data-stu-id="05232-169">Your Office Add-in can use additional UI features:</span></span>
  - [<span data-ttu-id="05232-170">Office アドインにカスタム キーボード ショートカットを追加する (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="05232-170">Add Custom keyboard shortcuts to your Office Add-ins (preview)</span></span>](../design/keyboard-shortcuts.md)
  - [<span data-ttu-id="05232-171">Office アドインでカスタム コンテキスト タブを作成する (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="05232-171">Create custom contextual tabs in Office Add-ins (preview)</span></span>](../design/contextual-tabs.md)
  - [<span data-ttu-id="05232-172">アドイン コマンドを有効または無効にする</span><span class="sxs-lookup"><span data-stu-id="05232-172">Enable and Disable Add-in Commands</span></span>](../design/disable-add-in-commands.md)
  - [<span data-ttu-id="05232-173">ドキュメントが開いたら、Office アドインでコードを実行する</span><span class="sxs-lookup"><span data-stu-id="05232-173">Run code in your Office Add-in when the document opens</span></span>](run-code-on-document-open.md)
  - [<span data-ttu-id="05232-174">Office アドインの作業ウィンドウを表示または非表示にする</span><span class="sxs-lookup"><span data-stu-id="05232-174">Show or hide the task pane of your Office Add-in</span></span>](show-hide-add-in.md)
- <span data-ttu-id="05232-175">Excel アドインの場合:</span><span class="sxs-lookup"><span data-stu-id="05232-175">For Excel add-ins:</span></span>
  - <span data-ttu-id="05232-176">カスタム関数で CORS がすべてサポートされます。</span><span class="sxs-lookup"><span data-stu-id="05232-176">Custom functions will have full CORS support.</span></span>
  - <span data-ttu-id="05232-177">カスタム関数で、Office.js API を呼び出して、スプレッドシート ドキュメントのデータを読み取ることができます。</span><span class="sxs-lookup"><span data-stu-id="05232-177">Custom functions can call Office.js APIs to read spreadsheet document data.</span></span>

<span data-ttu-id="05232-178">Windows 版 Office の場合、「[Office アドインで使用されるブラウザー](../concepts/browsers-used-by-office-web-add-ins.md)」で説明されているように、共有ランタイムには Microsoft Internet Explorer 11 ブラウザー インスタンスが必要です。また、アドインのリボンに表示するボタンはすべて、同じ共有ランタイムで実行されます。Office アドインで使用されるブラウザーで説明されているように、</span><span class="sxs-lookup"><span data-stu-id="05232-178">For Office on Windows, the shared runtime requires a Microsoft Internet Explorer 11 browser instance, as explained in [Browsers used by Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md). Additionally, any buttons that your add-in displays on the ribbon will run in the same shared runtime.</span></span> <span data-ttu-id="05232-179">次の図は、カスタム関数、リボン UI、作業ウィンドウのコードがすべて同じ JavaScript ランタイム内で実行される様子を示しています。</span><span class="sxs-lookup"><span data-stu-id="05232-179">The following image shows how custom functions, the ribbon UI, and the task pane code will all run in the same JavaScript runtime.</span></span>

![Excel の共有 IE ブラウザー ランタイムで実行されているカスタム関数、作業ウィンドウ、およびリボン ボタンの図](../images/custom-functions-in-browser-runtime.png)

### <a name="debugging"></a><span data-ttu-id="05232-181">デバッグ</span><span class="sxs-lookup"><span data-stu-id="05232-181">Debugging</span></span>

<span data-ttu-id="05232-182">共有ランタイムを使用している場合、この時点では、Windows の Excel でカスタム関数をデバッグするために Visual Studio Code を使用することはできません。</span><span class="sxs-lookup"><span data-stu-id="05232-182">When using a shared runtime, you can't use Visual Studio Code to debug custom functions in Excel on Windows at this time.</span></span> <span data-ttu-id="05232-183">代わりに開発者ツールを使用する必要があります。</span><span class="sxs-lookup"><span data-stu-id="05232-183">You'll need to use developer tools instead.</span></span> <span data-ttu-id="05232-184">さらに詳しい情報については、「[Windows 10 で開発者ツールを使用してアドインをデバッグする](../testing/debug-add-ins-using-f12-developer-tools-on-windows-10.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="05232-184">For more information, see [Debug add-ins using developer tools on Windows 10](../testing/debug-add-ins-using-f12-developer-tools-on-windows-10.md).</span></span>

### <a name="multiple-task-panes"></a><span data-ttu-id="05232-185">複数の作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="05232-185">Multiple task panes</span></span>

<span data-ttu-id="05232-186">共有ランタイムを使用する予定がある場合は、複数の作業ウィンドウを使用するようにアドインを設計しないでください。</span><span class="sxs-lookup"><span data-stu-id="05232-186">Don't design your add-in to use multiple task panes if you are planning to use a shared runtime.</span></span> <span data-ttu-id="05232-187">共有ランタイムは、1 つの作業ウィンドウのみサポートします。</span><span class="sxs-lookup"><span data-stu-id="05232-187">A shared runtime only supports the use of one task pane.</span></span> <span data-ttu-id="05232-188">`<TaskpaneID>` のない作業ウィンドウは、別の作業ウィンドウとして扱われますのでご注意ください。</span><span class="sxs-lookup"><span data-stu-id="05232-188">Note that any task pane without a `<TaskpaneID>` is considered a different task pane.</span></span>

## <a name="give-us-feedback"></a><span data-ttu-id="05232-189">ご意見ご感想をお寄せください</span><span class="sxs-lookup"><span data-stu-id="05232-189">Give us feedback</span></span>

<span data-ttu-id="05232-190">この機能について、ご意見をお待ちしております。</span><span class="sxs-lookup"><span data-stu-id="05232-190">We'd love to hear your feedback on this feature.</span></span> <span data-ttu-id="05232-191">バグや問題が発生したり、この機能について要求がございましたら、[office-js repo](https://github.com/OfficeDev/office-js) で GitHub に関する問題を作成してお知らせください。</span><span class="sxs-lookup"><span data-stu-id="05232-191">If you find any bugs, issues, or have requests on this feature, please let us know by creating a GitHub issue in the [office-js repo](https://github.com/OfficeDev/office-js).</span></span>

## <a name="see-also"></a><span data-ttu-id="05232-192">関連項目</span><span class="sxs-lookup"><span data-stu-id="05232-192">See also</span></span>

- [<span data-ttu-id="05232-193">カスタム関数から Excel API を呼び出す</span><span class="sxs-lookup"><span data-stu-id="05232-193">Call Excel APIs from a custom function</span></span>](../excel/call-excel-apis-from-custom-function.md)
- [<span data-ttu-id="05232-194">Office アドインにカスタム キーボード ショートカットを追加する (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="05232-194">Add custom keyboard shortcuts to your Office Add-ins (preview)</span></span>](../design/keyboard-shortcuts.md)
- [<span data-ttu-id="05232-195">Office アドインでカスタム コンテキスト タブを作成する (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="05232-195">Create custom contextual tabs in Office Add-ins (preview)</span></span>](../design/contextual-tabs.md)
- [<span data-ttu-id="05232-196">アドイン コマンドを有効または無効にする</span><span class="sxs-lookup"><span data-stu-id="05232-196">Enable and Disable Add-in Commands</span></span>](../design/disable-add-in-commands.md)
- [<span data-ttu-id="05232-197">ドキュメントが開いたら、Office アドインでコードを実行する</span><span class="sxs-lookup"><span data-stu-id="05232-197">Run code in your Office Add-in when the document opens</span></span>](run-code-on-document-open.md)
- [<span data-ttu-id="05232-198">Office アドインの作業ウィンドウを表示または非表示にする</span><span class="sxs-lookup"><span data-stu-id="05232-198">Show or hide the task pane of your Office Add-in</span></span>](show-hide-add-in.md)
- [<span data-ttu-id="05232-199">チュートリアル: Excel カスタム関数と作業ウィンドウの間でデータとイベントを共有する</span><span class="sxs-lookup"><span data-stu-id="05232-199">Tutorial: Share data and events between Excel custom functions and the task pane</span></span>](../tutorials/share-data-and-events-between-custom-functions-and-the-task-pane-tutorial.md)
