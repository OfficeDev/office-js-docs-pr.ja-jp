---
title: 'チュートリアル: Excel カスタム関数と作業ウィンドウの間でデータとイベントを共有する'
description: Excel でカスタム関数と作業ウィンドウの間でデータとイベントを共有する方法について説明します。
ms.date: 05/17/2020
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: a48d43270787648d8e5a53c885eab4b69cd8842e
ms.sourcegitcommit: cc6886b47c84ac37a3c957ff85dd0ed526ca5e43
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/12/2020
ms.locfileid: "46641152"
---
# <a name="tutorial-share-data-and-events-between-excel-custom-functions-and-the-task-pane"></a><span data-ttu-id="ee687-103">チュートリアル: Excel カスタム関数と作業ウィンドウの間でデータとイベントを共有する</span><span class="sxs-lookup"><span data-stu-id="ee687-103">Tutorial: Share data and events between Excel custom functions and the task pane</span></span>

<span data-ttu-id="ee687-104">共有ランタイムを使用するように Excel アドインを構成できます。</span><span class="sxs-lookup"><span data-stu-id="ee687-104">You can configure your Excel add-in to use a shared runtime.</span></span> <span data-ttu-id="ee687-105">これにより、グローバル データを共有したり、作業ウィンドウとユーザー設定の関数の間でイベントを送信したりできます。</span><span class="sxs-lookup"><span data-stu-id="ee687-105">This makes it possible to shared global data, or send events between the task pane and custom functions.</span></span>

<span data-ttu-id="ee687-106">ほとんどのカスタム関数のシナリオでは、作業ウィンドウのない（非表示の）カスタム関数を使用する特別な理由がない限り、共有ランタイムの使用をお勧めします。</span><span class="sxs-lookup"><span data-stu-id="ee687-106">For most custom functions scenarios, we recommend using a shared runtime, unless you have a specific reason to use a non-task pane (UI-less) custom function.</span></span>

<span data-ttu-id="ee687-107">このチュートリアルでは、Yo Office ジェネレーターを使用してアドイン プロジェクトを作成する方法に慣れていることを前提としています。</span><span class="sxs-lookup"><span data-stu-id="ee687-107">This tutorial assumes you're familiar with using the Yo Office generator to create add-in projects.</span></span> <span data-ttu-id="ee687-108">まだ使い慣れていない場合は、[Excel カスタム関数のチュートリアル](./excel-tutorial-create-custom-functions.md)を完了することを検討してください。</span><span class="sxs-lookup"><span data-stu-id="ee687-108">Consider completing the [Excel custom functions tutorial](./excel-tutorial-create-custom-functions.md), if you haven't already.</span></span>

## <a name="create-the-add-in-project"></a><span data-ttu-id="ee687-109">アドイン プロジェクトの作成</span><span class="sxs-lookup"><span data-stu-id="ee687-109">Create the add-in project</span></span>

<span data-ttu-id="ee687-110">Yeoman ジェネレーターを使用して、Excel アドイン プロジェクトを作成します。</span><span class="sxs-lookup"><span data-stu-id="ee687-110">Use the Yeoman generator to create an Excel add-in project.</span></span> <span data-ttu-id="ee687-111">次のコマンドを実行し、プロンプトに次の回答を入力します。</span><span class="sxs-lookup"><span data-stu-id="ee687-111">Run the following command and then answer the prompts with the following answers:</span></span>

```command line
yo office
```

- <span data-ttu-id="ee687-112">プロジェクトの種類を選択する:  **Excel カスタム関数アドイン プロジェクト**</span><span class="sxs-lookup"><span data-stu-id="ee687-112">Choose a project type: **Excel Custom Functions Add-in project**</span></span>
- <span data-ttu-id="ee687-113">スクリプトの種類を選択する:  **JavaScript**</span><span class="sxs-lookup"><span data-stu-id="ee687-113">Choose a script type: **JavaScript**</span></span>
- <span data-ttu-id="ee687-114">アドインの名前を何にしますか?  **個人用 Office アドイン**</span><span class="sxs-lookup"><span data-stu-id="ee687-114">What do you want to name your add-in? **My Office Add-in**</span></span>

![アドイン プロジェクトを作成するための Office からのプロンプトへ応答するスクリーンショット。](../images/yo-office-excel-project.png)

<span data-ttu-id="ee687-116">ウィザードを完了すると、ジェネレーターによってプロジェクトが作成され、サポートしているノード コンポーネントがインストールされます。</span><span class="sxs-lookup"><span data-stu-id="ee687-116">After you complete the wizard, the generator creates the project and installs supporting Node components.</span></span>

## <a name="configure-the-manifest"></a><span data-ttu-id="ee687-117">マニフェストを構成する</span><span class="sxs-lookup"><span data-stu-id="ee687-117">Configure the manifest</span></span>

1. <span data-ttu-id="ee687-118">Visual Studio Code を開始して [**個人用 Office アドイン**] プロジェクトを開きます。</span><span class="sxs-lookup"><span data-stu-id="ee687-118">Start Visual Studio Code and open the **My Office Add-in** project.</span></span>
2. <span data-ttu-id="ee687-119">
            **manifest.xml** ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="ee687-119">Open the **manifest.xml** file.</span></span>
3. <span data-ttu-id="ee687-120">`<VersionOverrides>` セクションを探し、次の `<Runtimes>` セクションを追加します。</span><span class="sxs-lookup"><span data-stu-id="ee687-120">Find the `<VersionOverrides>` section, and add the following `<Runtimes>` section.</span></span> <span data-ttu-id="ee687-121">作業ウィンドウを閉じてもカスタム関数が引き続き機能するように、有効期間は**長く**する必要があります。</span><span class="sxs-lookup"><span data-stu-id="ee687-121">The lifetime needs to be **long** so that the custom functions can still work even when the task pane is closed.</span></span>

   ```xml
   <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
     <Hosts>
       <Host xsi:type="Workbook">
         <Runtimes>
           <Runtime resid="ContosoAddin.Url" lifetime="long" />
         </Runtimes>
       <AllFormFactors>
   ```

4. <span data-ttu-id="ee687-122">`<Page>` 要素で、ソースの場所を **Functions.Page.Url** から **ContosoAddin.Url** に変更します。</span><span class="sxs-lookup"><span data-stu-id="ee687-122">In the `<Page>` element, change the source location from **Functions.Page.Url** to **ContosoAddin.Url**.</span></span>

   ```xml
   <AllFormFactors>
   ...
   <Page>
   <SourceLocation resid="ContosoAddin.Url"/>
   </Page>
   ...
   ```

5. <span data-ttu-id="ee687-123">`<DesktopFormFactor>` セクションで、**ContosoAddin.Url** を使用するように、**Command.Url** から **FunctionFile** を変更します。</span><span class="sxs-lookup"><span data-stu-id="ee687-123">In the `<DesktopFormFactor>` section, change the **FunctionFile** from **Commands.Url** to use **ContosoAddin.Url**.</span></span>

   ```xml
   <DesktopFormFactor>
   <GetStarted>
   ...
   </GetStarted>
   <FunctionFile resid="ContosoAddin.Url"/>
   ```

6. <span data-ttu-id="ee687-124">`<Action>` セクションで、ソースの場所を **Taskpane.Url** から **ContosoAddin.Url** に変更します。</span><span class="sxs-lookup"><span data-stu-id="ee687-124">In the `<Action>` section, change the source location from **Taskpane.Url** to **ContosoAddin.Url**.</span></span>

   ```xml
   <Action xsi:type="ShowTaskpane">
   <TaskpaneId>ButtonId1</TaskpaneId>
   <SourceLocation resid="ContosoAddin.Url"/>
   </Action>
   ```

7. <span data-ttu-id="ee687-125">**taskpane.html** を指す **ContosoAddin.Url** の新しい **Url ID** を追加します。</span><span class="sxs-lookup"><span data-stu-id="ee687-125">Add a new **Url id** for **ContosoAddin.Url** that points to **taskpane.html**.</span></span>

   ```xml
   <bt:Urls>
   <bt:Url id="Functions.Script.Url" DefaultValue="https://localhost:3000/dist/functions.js"/>
   ...
   <bt:Url id="ContosoAddin.Url" DefaultValue="https://localhost:3000/taskpane.html"/>
   ...
   ```

8. <span data-ttu-id="ee687-126">変更を保存してプロジェクトを再ビルドします。</span><span class="sxs-lookup"><span data-stu-id="ee687-126">Save your changes and rebuild the project.</span></span>

   ```command line
   npm run build
   ```

## <a name="share-state-between-custom-function-and-task-pane-code"></a><span data-ttu-id="ee687-127">カスタム関数と作業ウィンドウのコードの間で状態を共有する</span><span class="sxs-lookup"><span data-stu-id="ee687-127">Share state between custom function and task pane code</span></span>

<span data-ttu-id="ee687-128">カスタム関数が作業ウィンドウのコードと同じコンテキストで実行されるようになったため、**ストレージ** オブジェクトを使用せずに状態を直接共有できます。</span><span class="sxs-lookup"><span data-stu-id="ee687-128">Now that custom functions run in the same context as your task pane code, they can share state directly without using the **Storage** object.</span></span> <span data-ttu-id="ee687-129">次の手順は、カスタム関数と作業ウィンドウのコードの間でグローバル変数を共有する方法を示します。</span><span class="sxs-lookup"><span data-stu-id="ee687-129">The following instructions show how to share a global variable between custom function and task pane code.</span></span>

### <a name="create-custom-functions-to-get-or-store-shared-state"></a><span data-ttu-id="ee687-130">共有状態を取得または保存するカスタム関数を作成する</span><span class="sxs-lookup"><span data-stu-id="ee687-130">Create custom functions to get or store shared state</span></span>

1. <span data-ttu-id="ee687-131">Visual Studio Code でファイル **src/functions/functions.js** を開きます。</span><span class="sxs-lookup"><span data-stu-id="ee687-131">In Visual Studio Code open the file **src/functions/functions.js**.</span></span>
2. <span data-ttu-id="ee687-132">1 行目で、次のコードを一番上に挿入します。</span><span class="sxs-lookup"><span data-stu-id="ee687-132">On line 1, insert the following code at the very top.</span></span> <span data-ttu-id="ee687-133">これにより、**sharedState** という名前のグローバル変数が初期化されます。</span><span class="sxs-lookup"><span data-stu-id="ee687-133">This will initialize a global variable named **sharedState**.</span></span>

   ```js
   window.sharedState = "empty";
   ```

3. <span data-ttu-id="ee687-134">次のコードを追加して、値を **sharedState** 変数に保存するカスタム関数を作成します。</span><span class="sxs-lookup"><span data-stu-id="ee687-134">Add the following code to create a custom function that stores values to the **sharedState** variable.</span></span>

   ```js
   /**
    * Saves a string value to shared state with the task pane
    * @customfunction STOREVALUE
    * @param {string} value String to write to shared state with task pane.
    * @return {string} A success value
    */
   function storeValue(sharedValue) {
     window.sharedState = sharedValue;
     return "value stored";
   }
   ```

4. <span data-ttu-id="ee687-135">次のコードを追加して、**sharedState** 変数の現在の値を取得するカスタム関数を作成します。</span><span class="sxs-lookup"><span data-stu-id="ee687-135">Add the following code to create a custom function that gets the current value of the **sharedState** variable.</span></span>

   ```js
   /**
    * Gets a string value from shared state with the task pane
    * @customfunction GETVALUE
    * @returns {string} String value of the shared state with task pane.
    */
   function getValue() {
     return window.sharedState;
   }
   ```

5. <span data-ttu-id="ee687-136">ファイルを保存します。</span><span class="sxs-lookup"><span data-stu-id="ee687-136">Save the file.</span></span>

### <a name="create-task-pane-controls-to-work-with-global-data"></a><span data-ttu-id="ee687-137">グローバル データを操作する作業ウィンドウのコントロールを作成する</span><span class="sxs-lookup"><span data-stu-id="ee687-137">Create task pane controls to work with global data</span></span>

1. <span data-ttu-id="ee687-138">ファイル **src/taskpane/taskpane.html** を開きます。</span><span class="sxs-lookup"><span data-stu-id="ee687-138">Open the file **src/taskpane/taskpane.html**.</span></span>
2. <span data-ttu-id="ee687-139">`</head>` 要素の前に、次のスクリプト要素を追加します。</span><span class="sxs-lookup"><span data-stu-id="ee687-139">Add the following script element just before the `</head>` element.</span></span>

   ```html
   <script src="functions.js"></script>
   ```

3. <span data-ttu-id="ee687-140">終了 `</main>` 要素の後に、次の HTML を追加します。</span><span class="sxs-lookup"><span data-stu-id="ee687-140">After the closing `</main>` element, add the following HTML.</span></span> <span data-ttu-id="ee687-141">HTML は、グローバル データの取得または保存に使用される 2 つのテキスト ボックスとボタンを作成します。</span><span class="sxs-lookup"><span data-stu-id="ee687-141">The HTML creates two text boxes and buttons used to get or store global data.</span></span>

   ```html
   <ol>
     <li>
       Enter a value to send to the custom function and select
       <strong>Store</strong>.
     </li>
     <li>
       Enter <strong>=CONTOSO.GETVALUE()</strong>strong> into a cell to retrieve
       it.
     </li>
     <li>
       To send data to the task pane, in a cell, enter
       <strong>=CONTOSO.STOREVALUE("new value")</strong>
     </li>
     <li>Select <strong>Get</strong> to display the value in the task pane.</li>
   </ol>
   <p>Store new value to shared state</p>
   <div>
     <input type="text" id="storeBox" />
     <button onclick="storeSharedValue()">Store</button>
   </div>

   <p>Get shared state value</p>
   <div>
     <input type="text" id="getBox" />
     <button onclick="getSharedValue()">Get</button>
   </div>
   ```

4. <span data-ttu-id="ee687-142">`<body>` 要素の前に、次のスクリプトを追加します。</span><span class="sxs-lookup"><span data-stu-id="ee687-142">Before the `<body>` element add the following script.</span></span> <span data-ttu-id="ee687-143">このコードは、ユーザーがグローバル データを保存または取得するときにボタンのクリック イベントを処理します。</span><span class="sxs-lookup"><span data-stu-id="ee687-143">This code will handle the button click events when the user wants to store or get global data.</span></span>

   ```js
   <script>
   function storeSharedValue() {
   let sharedValue = document.getElementById('storeBox').value;
   window.sharedState = sharedValue;
   }

   function getSharedValue() {
   document.getElementById('getBox').value = window.sharedState;
   }</script>
   ```

5. <span data-ttu-id="ee687-144">ファイルを保存します。</span><span class="sxs-lookup"><span data-stu-id="ee687-144">Save the file.</span></span>
6. <span data-ttu-id="ee687-145">プロジェクトをビルドする</span><span class="sxs-lookup"><span data-stu-id="ee687-145">Build the project</span></span>

   ```command line
   npm run build
   ```

### <a name="try-sharing-data-between-the-custom-functions-and-task-pane"></a><span data-ttu-id="ee687-146">カスタム関数と作業ウィンドウの間でデータの共有を試す</span><span class="sxs-lookup"><span data-stu-id="ee687-146">Try sharing data between the custom functions and task pane</span></span>

- <span data-ttu-id="ee687-147">次のコマンドを使用してプロジェクトを開始します。</span><span class="sxs-lookup"><span data-stu-id="ee687-147">Start the project by using the following command.</span></span>

  ```command line
  npm run start
  ```

<span data-ttu-id="ee687-148">Excel が起動したら、作業ウィンドウのボタンを使用して共有データを保存または取得できます。</span><span class="sxs-lookup"><span data-stu-id="ee687-148">Once Excel starts, you can use the task pane buttons to store or get shared data.</span></span> <span data-ttu-id="ee687-149">カスタム関数のセルに `=CONTOSO.GETVALUE()` を入力して、同じ共有データを取得します。</span><span class="sxs-lookup"><span data-stu-id="ee687-149">Enter `=CONTOSO.GETVALUE()` into a cell for the custom function to retrieve the same shared data.</span></span> <span data-ttu-id="ee687-150">または `=CONTOSO.STOREVALUE("new value")` を使用して、共有データを新しい値に変更します。</span><span class="sxs-lookup"><span data-stu-id="ee687-150">Or use `=CONTOSO.STOREVALUE("new value")` to change the shared data to a new value.</span></span>

> [!NOTE]
> <span data-ttu-id="ee687-151">この記事で示すように、プロジェクトを構成すると、カスタム機能と作業ウィンドウのコンテキストが共有されます。</span><span class="sxs-lookup"><span data-stu-id="ee687-151">Configuring your project as shown in this article will share context between custom functions and the task pane.</span></span> <span data-ttu-id="ee687-152">カスタム関数から一部の Office API を呼び出すことができます。</span><span class="sxs-lookup"><span data-stu-id="ee687-152">Calling some Office APIs from custom functions is possible.</span></span> <span data-ttu-id="ee687-153">詳細については、「[カスタム関数からの Microsoft Excel API の呼び出しについて](../excel/call-excel-apis-from-custom-function.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="ee687-153">[See Call Microsoft Excel APIs from a custom function](../excel/call-excel-apis-from-custom-function.md) for more details.</span></span>
