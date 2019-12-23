---
ms.date: 11/04/2019
title: 'チュートリアル: Excel カスタム関数と作業ウィンドウの間でデータとイベントを共有する (プレビュー)'
ms.prod: excel
description: Excel でカスタム関数と作業ウィンドウの間でデータとイベントを共有します。
localization_priority: Priority
ms.openlocfilehash: 16affeb29bd5950198f81f85e44adaf812067829
ms.sourcegitcommit: 8c5c5a1bd3fe8b90f6253d9850e9352ed0b283ee
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/19/2019
ms.locfileid: "40814132"
---
# <a name="tutorial-share-data-and-events-between-excel-custom-functions-and-the-task-pane-preview"></a><span data-ttu-id="be6dd-103">チュートリアル: Excel カスタム関数と作業ウィンドウの間でデータとイベントを共有する (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="be6dd-103">Tutorial: Share data and events between Excel custom functions and the task pane (preview)</span></span>

<span data-ttu-id="be6dd-104">Excel カスタム関数と作業ウィンドウはグローバル データを共有し、互いに関数呼び出しを行うことができます。</span><span class="sxs-lookup"><span data-stu-id="be6dd-104">Excel custom functions and the task pane share global data, and can make function calls into each other.</span></span> <span data-ttu-id="be6dd-105">カスタム関数が作業ウィンドウで機能するようにプロジェクトを構成するには、この記事の指示に従ってください。</span><span class="sxs-lookup"><span data-stu-id="be6dd-105">To configure your project so that custom functions can work with the task pane, follow the instructions in this article.</span></span>

> [!NOTE]
> <span data-ttu-id="be6dd-106">この記事で説明する機能は現在プレビュー中であり、変更される可能性があります。</span><span class="sxs-lookup"><span data-stu-id="be6dd-106">The features described in this article are currently in preview and subject to change.</span></span> <span data-ttu-id="be6dd-107">これらを運用環境で使用することは現在サポートされていません。</span><span class="sxs-lookup"><span data-stu-id="be6dd-107">They are not currently supported for use in production environments.</span></span> <span data-ttu-id="be6dd-108">この記事のプレビュー機能は、Windows 上の Excel でのみ使用できます。</span><span class="sxs-lookup"><span data-stu-id="be6dd-108">The preview features in this article are only available on Excel on Windows.</span></span> <span data-ttu-id="be6dd-109">プレビュー機能を試すには、[Office Insider に参加する](https://insider.office.com/join)必要があります。</span><span class="sxs-lookup"><span data-stu-id="be6dd-109">To try the preview features, you will need to [join Office Insider](https://insider.office.com/join).</span></span>  <span data-ttu-id="be6dd-110">プレビュー機能を試す良い方法は、Office 365 サブスクリプションを使用することです。</span><span class="sxs-lookup"><span data-stu-id="be6dd-110">A good way to try out preview features is by using an Office 365 subscription.</span></span> <span data-ttu-id="be6dd-111">Office 365 サブスクリプションをまだお持ちでない場合は、[Office 365 Developer Program](https://developer.microsoft.com/office/dev-program) に参加することで入手できます。</span><span class="sxs-lookup"><span data-stu-id="be6dd-111">If you don't already have an Office 365 subscription, you can get one by joining the [Office 365 Developer Program](https://developer.microsoft.com/office/dev-program).</span></span>

## <a name="create-the-add-in-project"></a><span data-ttu-id="be6dd-112">アドイン プロジェクトの作成</span><span class="sxs-lookup"><span data-stu-id="be6dd-112">Create the add-in project</span></span>

<span data-ttu-id="be6dd-113">Yeoman ジェネレーターを使用して、Excel アドイン プロジェクトを作成します。</span><span class="sxs-lookup"><span data-stu-id="be6dd-113">Use the Yeoman generator to create an Excel add-in project.</span></span> <span data-ttu-id="be6dd-114">次のコマンドを実行し、プロンプトに次の回答を入力します。</span><span class="sxs-lookup"><span data-stu-id="be6dd-114">Run the following command and then answer the prompts with the following answers:</span></span>

```command&nbsp;line
yo office
```

- <span data-ttu-id="be6dd-115">プロジェクトの種類を選択する:  **Excel カスタム関数アドイン プロジェクト**</span><span class="sxs-lookup"><span data-stu-id="be6dd-115">Choose a project type: **Excel Custom Functions Add-in project**</span></span>
- <span data-ttu-id="be6dd-116">スクリプトの種類を選択する:  **JavaScript**</span><span class="sxs-lookup"><span data-stu-id="be6dd-116">Choose a script type: **JavaScript**</span></span>
- <span data-ttu-id="be6dd-117">アドインの名前を何にしますか?  **個人用 Office アドイン**</span><span class="sxs-lookup"><span data-stu-id="be6dd-117">What do you want to name your add-in? **My Office Add-in**</span></span>

![アドイン プロジェクトを作成するための Office からのプロンプトへ応答するスクリーンショット。](../images/yo-office-excel-project.png)

<span data-ttu-id="be6dd-119">ウィザードを完了すると、ジェネレーターによってプロジェクトが作成され、サポートしているノード コンポーネントがインストールされます。</span><span class="sxs-lookup"><span data-stu-id="be6dd-119">After you complete the wizard, the generator creates the project and installs supporting Node components.</span></span>

## <a name="configure-the-manifest"></a><span data-ttu-id="be6dd-120">マニフェストを構成する</span><span class="sxs-lookup"><span data-stu-id="be6dd-120">Configure the manifest</span></span>

1. <span data-ttu-id="be6dd-121">Visual Studio Code を開始して [**個人用 Office アドイン**] プロジェクトを開きます。</span><span class="sxs-lookup"><span data-stu-id="be6dd-121">Start Visual Studio Code and open the **My Office Add-in** project.</span></span>
2. <span data-ttu-id="be6dd-122">**manifest.xml** ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="be6dd-122">Open the **manifest.xml** file.</span></span>
3. <span data-ttu-id="be6dd-123">次のコードに示ように、**CustomFunctionsRuntime** バージョン **1.2** を使用するように `<Requirements>` ションを変更します。</span><span class="sxs-lookup"><span data-stu-id="be6dd-123">Change the `<Requirements>` section to use **CustomFunctionsRuntime** version **1.2** as shown in the following code.</span></span>
    
    ```xml
    <Requirements> 
    <Sets DefaultMinVersion="1.1">
    <Set Name="CustomFunctionsRuntime" MinVersion="1.2"/>
    </Sets>
    </Requirements>
    ```
    
4. <span data-ttu-id="be6dd-124">ブックの `<Host>` 要素の下に、次の `<Runtimes>` セクションを追加します。</span><span class="sxs-lookup"><span data-stu-id="be6dd-124">Under the `<Host>` element for the workbook, add the following `<Runtimes>` section.</span></span> <span data-ttu-id="be6dd-125">作業ウィンドウを閉じてもカスタム関数が引き続き機能するように、有効期間は**長く**する必要があります。</span><span class="sxs-lookup"><span data-stu-id="be6dd-125">The lifetime needs to be **long** so that the custom functions can still work even when the task pane is closed.</span></span>
    
    ```xml
    <Hosts>
    <Host xsi:type="Workbook">
    <Runtimes>
    <Runtime resid="TaskPaneAndCustomFunction.Url" lifetime="long" />
    </Runtimes>
    ```
    
5. <span data-ttu-id="be6dd-126">`<Page>` 要素で、ソースの場所を **Functions.Page.Url** から **TaskPaneAndCustomFunction.Url** に変更します。</span><span class="sxs-lookup"><span data-stu-id="be6dd-126">In the `<Page>` element, change the source location from **Functions.Page.Url** to **TaskPaneAndCustomFunction.Url**.</span></span>

    ```xml
    <AllFormFactors>
    ...
    <Page>
    <SourceLocation resid="TaskPaneAndCustomFunction.Url"/>
    </Page>
    ...
    ```

6. <span data-ttu-id="be6dd-127">`<DesktopFormFactor>` セクションで、**TaskPaneAndCustomFunction.Url** を使用するように、**Command.Url** から **FunctionFile** を変更します。</span><span class="sxs-lookup"><span data-stu-id="be6dd-127">In the `<DesktopFormFactor>` section, change the **FunctionFile** from **Commands.Url** to use **TaskPaneAndCustomFunction.Url**.</span></span>
    
    ```xml
    <DesktopFormFactor>
    <GetStarted>
    ...
    </GetStarted>
    <FunctionFile resid="TaskPaneAndCustomFunction.Url"/>
    ```
    
7. <span data-ttu-id="be6dd-128">`<Action>` セクションで、ソースの場所を **Taskpane.Url** から **TaskPaneAndCustomFunction.Url** に変更します。</span><span class="sxs-lookup"><span data-stu-id="be6dd-128">In the `<Action>` section, change the source location from **Taskpane.Url** to **TaskPaneAndCustomFunction.Url**.</span></span>
    
    ```xml
    <Action xsi:type="ShowTaskpane">
    <TaskpaneId>ButtonId1</TaskpaneId>
    <SourceLocation resid="TaskPaneAndCustomFunction.Url"/>
    </Action>
    ```
    
8. <span data-ttu-id="be6dd-129">**taskpane.html** を指す **TaskPaneAndCustomFunction.Url** の新しい **Url ID** を追加します。</span><span class="sxs-lookup"><span data-stu-id="be6dd-129">Add a new **Url id** for **TaskPaneAndCustomFunction.Url** that points to **taskpane.html**.</span></span>
     
    ```xml
    <bt:Urls>
    <bt:Url id="Functions.Script.Url" DefaultValue="https://localhost:3000/dist/functions.js"/>
    ...
    <bt:Url id="TaskPaneAndCustomFunction.Url" DefaultValue="https://localhost:3000/taskpane.html"/>
    ...
    ```
    
9. <span data-ttu-id="be6dd-130">変更を保存してプロジェクトを再ビルドします。</span><span class="sxs-lookup"><span data-stu-id="be6dd-130">Save your changes and rebuild the project.</span></span>
    
    ```command&nbsp;line
    npm run build
    ```

## <a name="share-state-between-custom-function-and-task-pane-code"></a><span data-ttu-id="be6dd-131">カスタム関数と作業ウィンドウのコードの間で状態を共有する</span><span class="sxs-lookup"><span data-stu-id="be6dd-131">Share state between custom function and task pane code</span></span> 

<span data-ttu-id="be6dd-132">カスタム関数が作業ウィンドウのコードと同じコンテキストで実行されるようになったため、**ストレージ** オブジェクトを使用せずに状態を直接共有できます。</span><span class="sxs-lookup"><span data-stu-id="be6dd-132">Now that custom functions run in the same context as your task pane code, they can share state directly without using the **Storage** object.</span></span> <span data-ttu-id="be6dd-133">次の手順は、カスタム関数と作業ウィンドウのコードの間でグローバル変数を共有する方法を示します。</span><span class="sxs-lookup"><span data-stu-id="be6dd-133">The following instructions show how to share a global variable between custom function and task pane code.</span></span>

### <a name="create-custom-functions-to-get-or-store-shared-state"></a><span data-ttu-id="be6dd-134">共有状態を取得または保存するカスタム関数を作成する</span><span class="sxs-lookup"><span data-stu-id="be6dd-134">Create custom functions to get or store shared state</span></span>

1. <span data-ttu-id="be6dd-135">Visual Studio Code でファイル **src/functions/functions.js** を開きます。</span><span class="sxs-lookup"><span data-stu-id="be6dd-135">In Visual Studio Code open the file **src/functions/functions.js**.</span></span>
2. <span data-ttu-id="be6dd-136">1 行目で、次のコードを一番上に挿入します。</span><span class="sxs-lookup"><span data-stu-id="be6dd-136">On line 1, insert the following code at the very top.</span></span> <span data-ttu-id="be6dd-137">これにより、**sharedState** という名前のグローバル変数が初期化されます。</span><span class="sxs-lookup"><span data-stu-id="be6dd-137">This will initialize a global variable named **sharedState**.</span></span>
    
    ```js
    window.sharedState = "empty";
    ```
    
3. <span data-ttu-id="be6dd-138">次のコードを追加して、値を **sharedState** 変数に保存するカスタム関数を作成します。</span><span class="sxs-lookup"><span data-stu-id="be6dd-138">Add the following code to create a custom function that stores values to the **sharedState** variable.</span></span>
    
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
    
4. <span data-ttu-id="be6dd-139">次のコードを追加して、**sharedState** 変数の現在の値を取得するカスタム関数を作成します。</span><span class="sxs-lookup"><span data-stu-id="be6dd-139">Add the following code to create a custom function that gets the current value of the **sharedState** variable.</span></span>

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
    
5. <span data-ttu-id="be6dd-140">ファイルを保存します。</span><span class="sxs-lookup"><span data-stu-id="be6dd-140">Save the file.</span></span>

### <a name="create-task-pane-controls-to-work-with-global-data"></a><span data-ttu-id="be6dd-141">グローバル データを操作する作業ウィンドウのコントロールを作成する</span><span class="sxs-lookup"><span data-stu-id="be6dd-141">Create task pane controls to work with global data</span></span> 

1. <span data-ttu-id="be6dd-142">ファイル **src/taskpane/taskpane.html** を開きます。</span><span class="sxs-lookup"><span data-stu-id="be6dd-142">Open the file**src/taskpane/taskpane.html**.</span></span>
2. <span data-ttu-id="be6dd-143">終了 `</main>` 要素の後に、次の HTML を追加します。</span><span class="sxs-lookup"><span data-stu-id="be6dd-143">After the closing `</main>` element, add the following HTML.</span></span> <span data-ttu-id="be6dd-144">HTML は、グローバル データの取得または保存に使用される 2 つのテキスト ボックスとボタンを作成します。</span><span class="sxs-lookup"><span data-stu-id="be6dd-144">The HTML creates two text boxes and buttons used to get or store global data.</span></span>

    ```html
    <ol>
    <li>Enter a value to send to the custom function and select <strong>Store</strong>.</li>
    <li>Enter <strong>=CONTOSO.GETVALUE()</strong>strong> into a cell to retrieve it.</li>
    <li>To send data to the task pane, in a cell, enter <strong>=CONTOSO.STOREVALUE("new value")</strong></li>
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
    
3. <span data-ttu-id="be6dd-145">`<body>` 要素の前に、次のスクリプトを追加します。</span><span class="sxs-lookup"><span data-stu-id="be6dd-145">Before the `<body>` element add the following script.</span></span> <span data-ttu-id="be6dd-146">このコードは、ユーザーがグローバル データを保存または取得するときにボタンのクリック イベントを処理します。</span><span class="sxs-lookup"><span data-stu-id="be6dd-146">This code will handle the button click events when the user wants to store or get global data.</span></span>
    
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
    
4. <span data-ttu-id="be6dd-147">ファイルを保存します。</span><span class="sxs-lookup"><span data-stu-id="be6dd-147">Save the file.</span></span>
5. <span data-ttu-id="be6dd-148">プロジェクトをビルドする</span><span class="sxs-lookup"><span data-stu-id="be6dd-148">Build the project</span></span>
    
    ```command&nbsp;line
    npm run build 
    ```

### <a name="try-sharing-data-between-the-custom-functions-and-task-pane"></a><span data-ttu-id="be6dd-149">カスタム関数と作業ウィンドウの間でデータの共有を試す</span><span class="sxs-lookup"><span data-stu-id="be6dd-149">Try sharing data between the custom functions and task pane</span></span>

- <span data-ttu-id="be6dd-150">次のコマンドを使用してプロジェクトを開始します。</span><span class="sxs-lookup"><span data-stu-id="be6dd-150">Start the project by using the following command.</span></span>

    ```command&nbsp;line
    npm run start
    ```

<span data-ttu-id="be6dd-151">Excel が起動したら、作業ウィンドウのボタンを使用して共有データを保存または取得できます。</span><span class="sxs-lookup"><span data-stu-id="be6dd-151">Once Excel starts, you can use the task pane buttons to store or get shared data.</span></span> <span data-ttu-id="be6dd-152">カスタム関数のセルに `=CONTOSO.GETVALUE()` を入力して、同じ共有データを取得します。</span><span class="sxs-lookup"><span data-stu-id="be6dd-152">Enter `=CONTOSO.GETVALUE()` into a cell for the custom function to retrieve the same shared data.</span></span> <span data-ttu-id="be6dd-153">または `=CONTOSO.STOREVALUE(“new value”)` を使用して、共有データを新しい値に変更します。</span><span class="sxs-lookup"><span data-stu-id="be6dd-153">Or use `=CONTOSO.STOREVALUE(“new value”)` to change the shared data to a new value.</span></span>

> [!NOTE]
> <span data-ttu-id="be6dd-154">この記事で示すように、プロジェクトを構成すると、カスタム機能と作業ウィンドウのコンテキストが共有されます。</span><span class="sxs-lookup"><span data-stu-id="be6dd-154">Configuring your project as shown in this article will share context between custom functions and the task pane.</span></span> <span data-ttu-id="be6dd-155">プレビューでカスタム関数から Office API を呼び出すことはできません。</span><span class="sxs-lookup"><span data-stu-id="be6dd-155">Calling Office APIs from custom functions is not supported.</span></span>

