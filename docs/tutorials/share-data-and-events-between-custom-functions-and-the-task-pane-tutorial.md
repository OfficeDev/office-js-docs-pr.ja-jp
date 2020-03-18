---
title: 'チュートリアル: Excel カスタム関数と作業ウィンドウの間でデータとイベントを共有する (プレビュー)'
description: Excel でカスタム関数と作業ウィンドウの間でデータとイベントを共有します。
ms.date: 02/20/2020
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: 34f2f1006a592c3ee7ab63fdc643648ca26cd01f
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/17/2020
ms.locfileid: "42719729"
---
# <a name="tutorial-share-data-and-events-between-excel-custom-functions-and-the-task-pane-preview"></a><span data-ttu-id="fa9a7-103">チュートリアル: Excel カスタム関数と作業ウィンドウの間でデータとイベントを共有する (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="fa9a7-103">Tutorial: Share data and events between Excel custom functions and the task pane (preview)</span></span>

[!include[Running custom functions in browser runtime note](../includes/excel-shared-runtime-preview-note.md)]

<span data-ttu-id="fa9a7-104">共有ランタイムを使用するように Excel アドインを構成できます。</span><span class="sxs-lookup"><span data-stu-id="fa9a7-104">You can configure your Excel add-in to use a shared runtime.</span></span> <span data-ttu-id="fa9a7-105">これにより、グローバル データを共有したり、作業ウィンドウとユーザー設定の関数の間でイベントを送信したりできます。</span><span class="sxs-lookup"><span data-stu-id="fa9a7-105">This will make it possible to shared global data, or send events between the task pane and custom functions.</span></span>

## <a name="create-the-add-in-project"></a><span data-ttu-id="fa9a7-106">アドイン プロジェクトの作成</span><span class="sxs-lookup"><span data-stu-id="fa9a7-106">Create the add-in project</span></span>

<span data-ttu-id="fa9a7-107">Yeoman ジェネレーターを使用して、Excel アドイン プロジェクトを作成します。</span><span class="sxs-lookup"><span data-stu-id="fa9a7-107">Use the Yeoman generator to create an Excel add-in project.</span></span> <span data-ttu-id="fa9a7-108">次のコマンドを実行し、プロンプトに次の回答を入力します。</span><span class="sxs-lookup"><span data-stu-id="fa9a7-108">Run the following command and then answer the prompts with the following answers:</span></span>

```command line
yo office
```

- <span data-ttu-id="fa9a7-109">プロジェクトの種類を選択する:  **Excel カスタム関数アドイン プロジェクト**</span><span class="sxs-lookup"><span data-stu-id="fa9a7-109">Choose a project type: **Excel Custom Functions Add-in project**</span></span>
- <span data-ttu-id="fa9a7-110">スクリプトの種類を選択する:  **JavaScript**</span><span class="sxs-lookup"><span data-stu-id="fa9a7-110">Choose a script type: **JavaScript**</span></span>
- <span data-ttu-id="fa9a7-111">アドインの名前を何にしますか?  **個人用 Office アドイン**</span><span class="sxs-lookup"><span data-stu-id="fa9a7-111">What do you want to name your add-in? **My Office Add-in**</span></span>

![アドイン プロジェクトを作成するための Office からのプロンプトへ応答するスクリーンショット。](../images/yo-office-excel-project.png)

<span data-ttu-id="fa9a7-113">ウィザードを完了すると、ジェネレーターによってプロジェクトが作成され、サポートしているノード コンポーネントがインストールされます。</span><span class="sxs-lookup"><span data-stu-id="fa9a7-113">After you complete the wizard, the generator creates the project and installs supporting Node components.</span></span>

## <a name="configure-the-manifest"></a><span data-ttu-id="fa9a7-114">マニフェストを構成する</span><span class="sxs-lookup"><span data-stu-id="fa9a7-114">Configure the manifest</span></span>

1. <span data-ttu-id="fa9a7-115">Visual Studio Code を開始して [**個人用 Office アドイン**] プロジェクトを開きます。</span><span class="sxs-lookup"><span data-stu-id="fa9a7-115">Start Visual Studio Code and open the **My Office Add-in** project.</span></span>
2. <span data-ttu-id="fa9a7-116">
            **manifest.xml** ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="fa9a7-116">Open the **manifest.xml** file.</span></span>
3. <span data-ttu-id="fa9a7-117">`<VersionOverrides>` セクションを探し、次の `<Runtimes>` セクションを追加します。</span><span class="sxs-lookup"><span data-stu-id="fa9a7-117">Find the `<VersionOverrides>` section, and add the following `<Runtimes>` section.</span></span> <span data-ttu-id="fa9a7-118">作業ウィンドウを閉じてもカスタム関数が引き続き機能するように、有効期間は**長く**する必要があります。</span><span class="sxs-lookup"><span data-stu-id="fa9a7-118">The lifetime needs to be **long** so that the custom functions can still work even when the task pane is closed.</span></span>

   ```xml
   <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
     <Hosts>
       <Host xsi:type="Workbook">
         <Runtimes>
           <Runtime resid="ContosoAddin.Url" lifetime="long" />
         </Runtimes>
       <AllFormFactors>
   ```

4. <span data-ttu-id="fa9a7-119">`<Page>` 要素で、ソースの場所を **Functions.Page.Url** から **ContosoAddin.Url** に変更します。</span><span class="sxs-lookup"><span data-stu-id="fa9a7-119">In the `<Page>` element, change the source location from **Functions.Page.Url** to **ContosoAddin.Url**.</span></span>

   ```xml
   <AllFormFactors>
   ...
   <Page>
   <SourceLocation resid="ContosoAddin.Url"/>
   </Page>
   ...
   ```

5. <span data-ttu-id="fa9a7-120">`<DesktopFormFactor>` セクションで、**ContosoAddin.Url** を使用するように、**Command.Url** から **FunctionFile** を変更します。</span><span class="sxs-lookup"><span data-stu-id="fa9a7-120">In the `<DesktopFormFactor>` section, change the **FunctionFile** from **Commands.Url** to use **ContosoAddin.Url**.</span></span>

   ```xml
   <DesktopFormFactor>
   <GetStarted>
   ...
   </GetStarted>
   <FunctionFile resid="ContosoAddin.Url"/>
   ```

6. <span data-ttu-id="fa9a7-121">`<Action>` セクションで、ソースの場所を **Taskpane.Url** から **ContosoAddin.Url** に変更します。</span><span class="sxs-lookup"><span data-stu-id="fa9a7-121">In the `<Action>` section, change the source location from **Taskpane.Url** to **ContosoAddin.Url**.</span></span>

   ```xml
   <Action xsi:type="ShowTaskpane">
   <TaskpaneId>ButtonId1</TaskpaneId>
   <SourceLocation resid="ContosoAddin.Url"/>
   </Action>
   ```

7. <span data-ttu-id="fa9a7-122">**taskpane.html** を指す **ContosoAddin.Url** の新しい **Url ID** を追加します。</span><span class="sxs-lookup"><span data-stu-id="fa9a7-122">Add a new **Url id** for **ContosoAddin.Url** that points to **taskpane.html**.</span></span>

   ```xml
   <bt:Urls>
   <bt:Url id="Functions.Script.Url" DefaultValue="https://localhost:3000/dist/functions.js"/>
   ...
   <bt:Url id="ContosoAddin.Url" DefaultValue="https://localhost:3000/taskpane.html"/>
   ...
   ```

8. <span data-ttu-id="fa9a7-123">変更を保存してプロジェクトを再ビルドします。</span><span class="sxs-lookup"><span data-stu-id="fa9a7-123">Save your changes and rebuild the project.</span></span>

   ```command line
   npm run build
   ```

## <a name="share-state-between-custom-function-and-task-pane-code"></a><span data-ttu-id="fa9a7-124">カスタム関数と作業ウィンドウのコードの間で状態を共有する</span><span class="sxs-lookup"><span data-stu-id="fa9a7-124">Share state between custom function and task pane code</span></span>

<span data-ttu-id="fa9a7-125">カスタム関数が作業ウィンドウのコードと同じコンテキストで実行されるようになったため、**ストレージ** オブジェクトを使用せずに状態を直接共有できます。</span><span class="sxs-lookup"><span data-stu-id="fa9a7-125">Now that custom functions run in the same context as your task pane code, they can share state directly without using the **Storage** object.</span></span> <span data-ttu-id="fa9a7-126">次の手順は、カスタム関数と作業ウィンドウのコードの間でグローバル変数を共有する方法を示します。</span><span class="sxs-lookup"><span data-stu-id="fa9a7-126">The following instructions show how to share a global variable between custom function and task pane code.</span></span>

### <a name="create-custom-functions-to-get-or-store-shared-state"></a><span data-ttu-id="fa9a7-127">共有状態を取得または保存するカスタム関数を作成する</span><span class="sxs-lookup"><span data-stu-id="fa9a7-127">Create custom functions to get or store shared state</span></span>

1. <span data-ttu-id="fa9a7-128">Visual Studio Code でファイル **src/functions/functions.js** を開きます。</span><span class="sxs-lookup"><span data-stu-id="fa9a7-128">In Visual Studio Code open the file **src/functions/functions.js**.</span></span>
2. <span data-ttu-id="fa9a7-129">1 行目で、次のコードを一番上に挿入します。</span><span class="sxs-lookup"><span data-stu-id="fa9a7-129">On line 1, insert the following code at the very top.</span></span> <span data-ttu-id="fa9a7-130">これにより、**sharedState** という名前のグローバル変数が初期化されます。</span><span class="sxs-lookup"><span data-stu-id="fa9a7-130">This will initialize a global variable named **sharedState**.</span></span>

   ```js
   window.sharedState = "empty";
   ```

3. <span data-ttu-id="fa9a7-131">次のコードを追加して、値を **sharedState** 変数に保存するカスタム関数を作成します。</span><span class="sxs-lookup"><span data-stu-id="fa9a7-131">Add the following code to create a custom function that stores values to the **sharedState** variable.</span></span>

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

4. <span data-ttu-id="fa9a7-132">次のコードを追加して、**sharedState** 変数の現在の値を取得するカスタム関数を作成します。</span><span class="sxs-lookup"><span data-stu-id="fa9a7-132">Add the following code to create a custom function that gets the current value of the **sharedState** variable.</span></span>

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

5. <span data-ttu-id="fa9a7-133">ファイルを保存します。</span><span class="sxs-lookup"><span data-stu-id="fa9a7-133">Save the file.</span></span>

### <a name="create-task-pane-controls-to-work-with-global-data"></a><span data-ttu-id="fa9a7-134">グローバル データを操作する作業ウィンドウのコントロールを作成する</span><span class="sxs-lookup"><span data-stu-id="fa9a7-134">Create task pane controls to work with global data</span></span>

1. <span data-ttu-id="fa9a7-135">ファイル **src/taskpane/taskpane.html** を開きます。</span><span class="sxs-lookup"><span data-stu-id="fa9a7-135">Open the file **src/taskpane/taskpane.html**.</span></span>
2. <span data-ttu-id="fa9a7-136">`</head>` 要素の前に、次のスクリプト要素を追加します。</span><span class="sxs-lookup"><span data-stu-id="fa9a7-136">Add the following script element just before the `</head>` element.</span></span>

   ```html
   <script src="functions.js"></script>
   ```

3. <span data-ttu-id="fa9a7-137">終了 `</main>` 要素の後に、次の HTML を追加します。</span><span class="sxs-lookup"><span data-stu-id="fa9a7-137">After the closing `</main>` element, add the following HTML.</span></span> <span data-ttu-id="fa9a7-138">HTML は、グローバル データの取得または保存に使用される 2 つのテキスト ボックスとボタンを作成します。</span><span class="sxs-lookup"><span data-stu-id="fa9a7-138">The HTML creates two text boxes and buttons used to get or store global data.</span></span>

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

4. <span data-ttu-id="fa9a7-139">`<body>` 要素の前に、次のスクリプトを追加します。</span><span class="sxs-lookup"><span data-stu-id="fa9a7-139">Before the `<body>` element add the following script.</span></span> <span data-ttu-id="fa9a7-140">このコードは、ユーザーがグローバル データを保存または取得するときにボタンのクリック イベントを処理します。</span><span class="sxs-lookup"><span data-stu-id="fa9a7-140">This code will handle the button click events when the user wants to store or get global data.</span></span>

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

5. <span data-ttu-id="fa9a7-141">ファイルを保存します。</span><span class="sxs-lookup"><span data-stu-id="fa9a7-141">Save the file.</span></span>
6. <span data-ttu-id="fa9a7-142">プロジェクトをビルドする</span><span class="sxs-lookup"><span data-stu-id="fa9a7-142">Build the project</span></span>

   ```command line
   npm run build
   ```

### <a name="try-sharing-data-between-the-custom-functions-and-task-pane"></a><span data-ttu-id="fa9a7-143">カスタム関数と作業ウィンドウの間でデータの共有を試す</span><span class="sxs-lookup"><span data-stu-id="fa9a7-143">Try sharing data between the custom functions and task pane</span></span>

- <span data-ttu-id="fa9a7-144">次のコマンドを使用してプロジェクトを開始します。</span><span class="sxs-lookup"><span data-stu-id="fa9a7-144">Start the project by using the following command.</span></span>

  ```command line
  npm run start
  ```

<span data-ttu-id="fa9a7-145">Excel が起動したら、作業ウィンドウのボタンを使用して共有データを保存または取得できます。</span><span class="sxs-lookup"><span data-stu-id="fa9a7-145">Once Excel starts, you can use the task pane buttons to store or get shared data.</span></span> <span data-ttu-id="fa9a7-146">カスタム関数のセルに `=CONTOSO.GETVALUE()` を入力して、同じ共有データを取得します。</span><span class="sxs-lookup"><span data-stu-id="fa9a7-146">Enter `=CONTOSO.GETVALUE()` into a cell for the custom function to retrieve the same shared data.</span></span> <span data-ttu-id="fa9a7-147">または `=CONTOSO.STOREVALUE("new value")` を使用して、共有データを新しい値に変更します。</span><span class="sxs-lookup"><span data-stu-id="fa9a7-147">Or use `=CONTOSO.STOREVALUE("new value")` to change the shared data to a new value.</span></span>

> [!NOTE]
> <span data-ttu-id="fa9a7-148">この記事で示すように、プロジェクトを構成すると、カスタム機能と作業ウィンドウのコンテキストが共有されます。</span><span class="sxs-lookup"><span data-stu-id="fa9a7-148">Configuring your project as shown in this article will share context between custom functions and the task pane.</span></span> <span data-ttu-id="fa9a7-149">プレビューでカスタム関数から Office API を呼び出すことはできません。</span><span class="sxs-lookup"><span data-stu-id="fa9a7-149">Calling Office APIs from custom functions is not supported in the preview.</span></span>
