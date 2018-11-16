<span data-ttu-id="d62fb-101">チュートリアルのこの手順では、リボンに別のボタンを追加します。このボタンをクリックすると、ワークシートの保護のオン/オフが切り替わるように定義した関数が実行されるようにします。</span><span class="sxs-lookup"><span data-stu-id="d62fb-101">In this step of the tutorial, you'll add another button to the ribbon that, when chosen, executes a function that you'll define to toggle worksheet protection on and off.</span></span>

> [!NOTE]
> <span data-ttu-id="d62fb-102">このページでは、Excel のアドインのチュートリアルの個々 の手順について説明します。</span><span class="sxs-lookup"><span data-stu-id="d62fb-102">This page describes an individual step of the Excel add-in tutorial.</span></span> <span data-ttu-id="d62fb-103">このページに検索エンジンの結果から、または直接リンクからアクセスした場合は、「[Excel アドインのチュートリアル](../tutorials/excel-tutorial.yml)」の紹介ページに移動し、チュートリアルを最初から始めてください。</span><span class="sxs-lookup"><span data-stu-id="d62fb-103">If you’ve arrived at this page via search engine results or other direct link, please go to the [Excel add-in tutorial](../tutorials/excel-tutorial.yml) introduction page to start the tutorial from the beginning.</span></span>

## <a name="configure-the-manifest-to-add-a-second-ribbon-button"></a><span data-ttu-id="d62fb-104">2 つ目のリボン ボタンを追加するようにマニフェストを構成する</span><span class="sxs-lookup"><span data-stu-id="d62fb-104">Configure the manifest to add a second ribbon button</span></span>

1. <span data-ttu-id="d62fb-105">マニフェスト ファイル **my-office-add-in-manifest.xml** を開きます。</span><span class="sxs-lookup"><span data-stu-id="d62fb-105">Open the manifest file **my-office-add-in-manifest.xml**.</span></span>
2. <span data-ttu-id="d62fb-106">`<Control>` 要素を検索します。</span><span class="sxs-lookup"><span data-stu-id="d62fb-106">Find the `<Control>` element.</span></span> <span data-ttu-id="d62fb-107">この要素では、アドインの起動に使用している **[ホーム]** リボンの **[作業ウィンドウの表示]** ボタンを定義しています。</span><span class="sxs-lookup"><span data-stu-id="d62fb-107">This element defines the **Show Taskpane** button on the **Home** ribbon you have been using to launch the add-in.</span></span> <span data-ttu-id="d62fb-108">ここでは、**[ホーム]** リボンの同じグループに 2 つ目のボタンを追加します。</span><span class="sxs-lookup"><span data-stu-id="d62fb-108">We're going to add a second button to the same group on the **Home** ribbon.</span></span> <span data-ttu-id="d62fb-109">Control 終了タグ (`</Control>`) と Group 終了タグ (`</Group>`) の間に、次のマークアップを追加します。</span><span class="sxs-lookup"><span data-stu-id="d62fb-109">In between the end Control tag (`</Control>`) and the end Group tag (`</Group>`), add the following markup.</span></span>

    ```xml
    <Control xsi:type="Button" id="<!--TODO1: Unique (in manifest) name for button -->">
        <Label resid="<!--TODO2: Button label -->" />
        <Supertip>            
            <Title resid="<!-- TODO3: Button tool tip title -->" />
            <Description resid="<!-- TODO4: Button tool tip description -->" />
        </Supertip>
        <Icon>
            <bt:Image size="16" resid="Contoso.tpicon_16x16" />
            <bt:Image size="32" resid="Contoso.tpicon_32x32" />
            <bt:Image size="80" resid="Contoso.tpicon_80x80" />
        </Icon>
        <Action xsi:type="<!-- TODO5: Specify the type of action-->">
            <!-- TODO6: Identify the function.-->
        </Action>
    </Control>
    ```

3. <span data-ttu-id="d62fb-110">`TODO1` は文字列に置き換えて、このマニフェスト ファイル内で一意の ID をボタンに割り当てます。</span><span class="sxs-lookup"><span data-stu-id="d62fb-110">Replace `TODO1` with a string that gives the button an ID that is unique within this manifest file.</span></span> <span data-ttu-id="d62fb-111">このマニフェストには、別のボタンが 1 つしか存在していないため難しいことではありません。</span><span class="sxs-lookup"><span data-stu-id="d62fb-111">There's only one other button in the manifest, so this isn't difficult.</span></span> <span data-ttu-id="d62fb-112">このボタンでは、ワークシートの保護のオン/オフを切り替える予定なので、"ToggleProtection" を使用することにします。</span><span class="sxs-lookup"><span data-stu-id="d62fb-112">Since our button is going to toggle protection of the worksheet on and off, use "ToggleProtection".</span></span> <span data-ttu-id="d62fb-113">作業が完了すると、Control 開始タグの全体は次のようになります。</span><span class="sxs-lookup"><span data-stu-id="d62fb-113">When you are done, the entire start Control tag should look like the following:</span></span>

    ```xml
    <Control xsi:type="Button" id="ToggleProtection">
    ```

4. <span data-ttu-id="d62fb-114">その次の 3 つの `TODO` では、"resid" を設定します ("resid" はリソース ID の略号です)。</span><span class="sxs-lookup"><span data-stu-id="d62fb-114">The next three `TODO`s set "resid"s, which is short for resource ID.</span></span> <span data-ttu-id="d62fb-115">リソースは文字列です。これら 3 つの文字列は、この後の手順で作成します。</span><span class="sxs-lookup"><span data-stu-id="d62fb-115">A resource is a string, and you'll create these three strings in a later step.</span></span> <span data-ttu-id="d62fb-116">ここでは、そのリソースに ID を割り当てる必要があります。</span><span class="sxs-lookup"><span data-stu-id="d62fb-116">For now, you need to give IDs to the resources.</span></span> <span data-ttu-id="d62fb-117">ボタンのラベルは "Toggle Protection" と表示されるようにしますが、この文字列の *ID* は "ProtectionButtonLabel" にします。そのため、完成した `Label` 要素は次のコードのようになります。</span><span class="sxs-lookup"><span data-stu-id="d62fb-117">The button label should read "Toggle Protection", but the *ID* of this string should be "ProtectionButtonLabel", so the completed `Label` element should look like the following code:</span></span>

    ```xml
    <Label resid="ProtectionButtonLabel" />
    ```

5. <span data-ttu-id="d62fb-118">`SuperTip` 要素では、このボタンのツール ヒントを定義します。</span><span class="sxs-lookup"><span data-stu-id="d62fb-118">The `SuperTip` element defines the tool tip for the button.</span></span> <span data-ttu-id="d62fb-119">ツール ヒントのタイトルはボタンのラベルと同じにする必要があるため、リソース ID にはまったく同じ "ProtectionButtonLabel" を使用することにします。</span><span class="sxs-lookup"><span data-stu-id="d62fb-119">The tool tip title should be the same as the button label, so we use the very same resource ID: "ProtectionButtonLabel".</span></span> <span data-ttu-id="d62fb-120">ツール ヒントの説明は、"Click to turn protection of the worksheet on and off" にする予定です。</span><span class="sxs-lookup"><span data-stu-id="d62fb-120">The tool tip description will be "Click to turn protection of the worksheet on and off".</span></span> <span data-ttu-id="d62fb-121">ただし、`ID` は "ProtectionButtonToolTip" にします。</span><span class="sxs-lookup"><span data-stu-id="d62fb-121">But the `ID` should be "ProtectionButtonToolTip".</span></span> <span data-ttu-id="d62fb-122">作業が完了すると、`SuperTip` マークアップの全体は次のコードのようになります。</span><span class="sxs-lookup"><span data-stu-id="d62fb-122">So, when you are done, the whole `SuperTip` markup should look like the following code:</span></span> 

    ```xml
    <Supertip>            
        <Title resid="ProtectionButtonLabel" />
        <Description resid="ProtectionButtonToolTip" />
    </Supertip>
    ```

   > [!NOTE] 
   > <span data-ttu-id="d62fb-123">運用アドインでは、異なる 2 つのボタンに同じアイコンを使用することは避けたいところですが、このチュートリアルでは説明を簡単にするために同じアイコンを使用します。</span><span class="sxs-lookup"><span data-stu-id="d62fb-123">In a production add-in, you would not want to use the same icon for two different buttons; but to simplify this tutorial, we'll do that.</span></span> <span data-ttu-id="d62fb-124">そのため、この新しい `Control` の `Icon` マークアップは、単に既存の `Control` から `Icon` 要素をコピーします。</span><span class="sxs-lookup"><span data-stu-id="d62fb-124">So the `Icon` markup in our new `Control` is just a copy of the `Icon` element from the existing `Control`.</span></span> 

6. <span data-ttu-id="d62fb-125">既にマニフェストに存在している元の `Control` 要素の内側にある `Action` 要素では、その要素のタイプが `ShowTaskpane` に設定されていますが、新しいボタンで作業ウィンドウを開く予定はありません。このボタンでは、この後の手順で作成するカスタム関数を実行する予定です。</span><span class="sxs-lookup"><span data-stu-id="d62fb-125">The `Action` element inside the original `Control` element that was already present in the manifest, has its type set to `ShowTaskpane`, but our new button isn't going to open a task pane; it's going to run a custom function that you create in a later step.</span></span> <span data-ttu-id="d62fb-126">そのため、`TODO5` は、カスタム関数をトリガーするボタンのアクション タイプである `ExecuteFunction` に置き換えます。</span><span class="sxs-lookup"><span data-stu-id="d62fb-126">So replace `TODO5` with `ExecuteFunction` which is the action type for buttons that trigger custom functions.</span></span> <span data-ttu-id="d62fb-127">`Action` 開始タグは次のようになります。</span><span class="sxs-lookup"><span data-stu-id="d62fb-127">The start `Action` tag should look like the following code:</span></span>
 
    ```xml
    <Action xsi:type="ExecuteFunction">
    ```

7. <span data-ttu-id="d62fb-128">元の `Action` 要素には、作業ウィンドウ ID を指定する子要素と、作業ウィンドウで開かれるページの URL を指定する子要素があります。</span><span class="sxs-lookup"><span data-stu-id="d62fb-128">The original `Action` element has child elements that specify a task pane ID and a URL of the page that should be opened in the task pane.</span></span> <span data-ttu-id="d62fb-129">ただし、`ExecuteFunction` タイプの `Action` 要素には、実行を制御する関数の名前を指定する子要素を 1 つ含めます。</span><span class="sxs-lookup"><span data-stu-id="d62fb-129">But an `Action` element of the `ExecuteFunction` type has a single child element that names the function that the control executes.</span></span> <span data-ttu-id="d62fb-130">その関数は、`toggleProtection` という名前にして、この後の手順で作成します。</span><span class="sxs-lookup"><span data-stu-id="d62fb-130">You'll create that function in a later step, and it will be called `toggleProtection`.</span></span> <span data-ttu-id="d62fb-131">そのために、`TODO6` を次のマークアップに置き換えます。</span><span class="sxs-lookup"><span data-stu-id="d62fb-131">So, replace `TODO6` with the following markup:</span></span>
 
    ```xml
    <FunctionName>toggleProtection</FunctionName>
    ```

    <span data-ttu-id="d62fb-132">`Control` マークアップの全体は、次のようになりました。</span><span class="sxs-lookup"><span data-stu-id="d62fb-132">The entire `Control` markup should now look like the following:</span></span>

    ```xml
    <Control xsi:type="Button" id="ToggleProtection">
        <Label resid="ProtectionButtonLabel" />
        <Supertip>            
            <Title resid="ProtectionButtonLabel" />
            <Description resid="ProtectionButtonToolTip" />
        </Supertip>
        <Icon>
            <bt:Image size="16" resid="Contoso.tpicon_16x16" />
            <bt:Image size="32" resid="Contoso.tpicon_32x32" />
            <bt:Image size="80" resid="Contoso.tpicon_80x80" />
        </Icon>
        <Action xsi:type="ExecuteFunction">
           <FunctionName>toggleProtection</FunctionName>
        </Action>
    </Control>
    ```

8. <span data-ttu-id="d62fb-133">マニフェストの `Resources` セクションまで下にスクロールします。</span><span class="sxs-lookup"><span data-stu-id="d62fb-133">Scroll down to the `Resources` section of the manifest.</span></span>

9. <span data-ttu-id="d62fb-134">`bt:ShortStrings` 要素の子として、次のマークアップを追加します。</span><span class="sxs-lookup"><span data-stu-id="d62fb-134">Add the following markup as a child of the `bt:ShortStrings` element.</span></span>

    ```xml
    <bt:String id="ProtectionButtonLabel" DefaultValue="Toggle Worksheet Protection" />
    ```

10. <span data-ttu-id="d62fb-135">`bt:LongStrings` 要素の子として、次のマークアップを追加します。</span><span class="sxs-lookup"><span data-stu-id="d62fb-135">Add the following markup as a child of the `bt:LongStrings` element.</span></span>

    ```xml
    <bt:String id="ProtectionButtonToolTip" DefaultValue="Click to protect or unprotect the current worksheet." />
    ```

11. <span data-ttu-id="d62fb-136">ファイルを保存します。</span><span class="sxs-lookup"><span data-stu-id="d62fb-136">Be sure to save the file.</span></span>

## <a name="create-the-function-that-protects-the-sheet"></a><span data-ttu-id="d62fb-137">シートを保護する関数を作成する</span><span class="sxs-lookup"><span data-stu-id="d62fb-137">Create the function that protects the sheet</span></span>

1. <span data-ttu-id="d62fb-138">ファイル \function-file\function-file.js を開きます。</span><span class="sxs-lookup"><span data-stu-id="d62fb-138">Open the file \function-file\function-file.js.</span></span>

2. <span data-ttu-id="d62fb-139">このファイルには、即時実行関数式 (IIFE) が既に含まれています。</span><span class="sxs-lookup"><span data-stu-id="d62fb-139">The file already has an Immediately Invoked Function Expression (IFFE).</span></span> <span data-ttu-id="d62fb-140">カスタムの初期化ロジックは必要ないため、`Office.initialize` に割り当てられた関数は空のままにしておきます </span><span class="sxs-lookup"><span data-stu-id="d62fb-140">No custom initialization logic is needed, so leave the function that is assigned to `Office.initialize` with an empty body.</span></span> <span data-ttu-id="d62fb-141">(ただし、削除してはいけません。</span><span class="sxs-lookup"><span data-stu-id="d62fb-141">(But do not delete it.</span></span> <span data-ttu-id="d62fb-142">`Office.initialize` プロパティは Null や未定義にすることはできません)。*IIFE の外側に*、次のコードを追加します。</span><span class="sxs-lookup"><span data-stu-id="d62fb-142">The `Office.initialize` property cannot be null or undefined.) *Outside of the IIFE*, add the following code.</span></span> <span data-ttu-id="d62fb-143">メソッドに `args` パラメーターを指定していることと、メソッドの最後のほうの行で `args.completed` を呼び出していることに注目してください。</span><span class="sxs-lookup"><span data-stu-id="d62fb-143">Note that we specify an `args` parameter to the method and the very last line of the method calls `args.completed`.</span></span> <span data-ttu-id="d62fb-144">**ExecuteFunction** タイプのすべてのアドイン コマンドでは、これが要件になります。</span><span class="sxs-lookup"><span data-stu-id="d62fb-144">This is a requirement for all add-in commands of type **ExecuteFunction**.</span></span> <span data-ttu-id="d62fb-145">これにより、関数が終了したことと、UI が再度応答可能になることを Office ホスト アプリケーションに通知します。</span><span class="sxs-lookup"><span data-stu-id="d62fb-145">It signals the Office host application that the function has finished and the UI can become responsive again.</span></span>

    ```javascript
    function toggleProtection(args) {
        Excel.run(function (context) {
            
            // TODO1: Queue commands to reverse the protection status of the current worksheet.

            return context.sync();
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
        args.completed();
    }
    ```

3. <span data-ttu-id="d62fb-146">`TODO1` を次のコードに置き換えます。</span><span class="sxs-lookup"><span data-stu-id="d62fb-146">Replace `TODO1` with the following code.</span></span> <span data-ttu-id="d62fb-147">このコードでは、標準の切り替えパターンで、ワークシート オブジェクトの protection プロパティを使用します。</span><span class="sxs-lookup"><span data-stu-id="d62fb-147">This code uses the worksheet object's protection property in a standard toggle pattern.</span></span> <span data-ttu-id="d62fb-148">`TODO2` については、次のセクションで説明します。</span><span class="sxs-lookup"><span data-stu-id="d62fb-148">The `TODO2` will be explained in the next section.</span></span>

    ```javascript
    const sheet = context.workbook.worksheets.getActiveWorksheet();

    // TODO2: Queue command to load the sheet's "protection.protected" property from
    //        the document and re-synchronize the document and task pane.

     if (sheet.protection.protected) {
        sheet.protection.unprotect();
    } else {
        sheet.protection.protect();
    }
    ``` 

## <a name="add-code-to-fetch-document-properties-into-the-task-panes-script-objects"></a><span data-ttu-id="d62fb-149">ドキュメントのプロパティを作業ウィンドウのスクリプト オブジェクトにフェッチするコードを追加する</span><span class="sxs-lookup"><span data-stu-id="d62fb-149">Add code to fetch document properties into the task pane's script objects</span></span>

<span data-ttu-id="d62fb-150">このチュートリアルのシリーズで前述したすべての関数では、Office ドキュメントへの*書き込み*コマンドをキューに登録していました。</span><span class="sxs-lookup"><span data-stu-id="d62fb-150">In all the earlier functions in this series of tutorials, you queued commands to *write* to the Office document.</span></span> <span data-ttu-id="d62fb-151">各関数は、キューに登録されたコマンドを実行対象のドキュメントに送信する `context.sync()` メソッドを呼び出すことで終了しています。</span><span class="sxs-lookup"><span data-stu-id="d62fb-151">Each function ended with a call to the `context.sync()` method which sends the queued commands to the document to be executed.</span></span> <span data-ttu-id="d62fb-152">ただし、最後の手順で追加したコードでは、`sheet.protection.protected` プロパティを呼び出しています。このことが、これまでに作成した関数とは大きく異なります。`sheet` オブジェクトは、この作業ウィンドウのスクリプトに存在する単なるプロキシ オブジェクトなので、</span><span class="sxs-lookup"><span data-stu-id="d62fb-152">But the code you added in the last step calls the `sheet.protection.protected` property, and this is a significant difference from the earlier functions you wrote, because the `sheet` object is only a proxy object that exists in your task pane's script.</span></span> <span data-ttu-id="d62fb-153">ドキュメントの実際の保護の状態を認識できません。そのため、その `protection.protected` プロパティでは実際の値が保持できません。</span><span class="sxs-lookup"><span data-stu-id="d62fb-153">It doesn't know what the actual protection state of the document is, so its `protection.protected` property can't have a real value.</span></span> <span data-ttu-id="d62fb-154">まず、ドキュメントから保護の状態をフェッチする必要があり、その状態を使用して `sheet.protection.protected` の値を設定します。</span><span class="sxs-lookup"><span data-stu-id="d62fb-154">It is necessary to first fetch the protection status from the document and use it set the value of `sheet.protection.protected`.</span></span> <span data-ttu-id="d62fb-155">そのようにした場合にのみ、例外がスローされることなく `sheet.protection.protected` を呼び出せるようになります。</span><span class="sxs-lookup"><span data-stu-id="d62fb-155">Only then can `sheet.protection.protected` be called without causing an exception to be thrown.</span></span> <span data-ttu-id="d62fb-156">このフェッチ処理には、3 つの手順があります。</span><span class="sxs-lookup"><span data-stu-id="d62fb-156">This fetching process has three steps:</span></span>

   1. <span data-ttu-id="d62fb-157">コードで読み取る必要があるプロパティをロードする (つまりフェッチする) コマンドをキューに登録します。</span><span class="sxs-lookup"><span data-stu-id="d62fb-157">Queue a command to load (that is; fetch) the properties that your code needs to read.</span></span>
   2. <span data-ttu-id="d62fb-158">コンテキスト オブジェクトの `sync` メソッドを呼び出します。このメソッドは、キューに登録されたコマンドを実行対象のドキュメントに送信して、要求された情報を返します。</span><span class="sxs-lookup"><span data-stu-id="d62fb-158">Call the context object's `sync` method to send the queued command to the document for execution and return the requested information.</span></span>
   3. <span data-ttu-id="d62fb-159">`sync` メソッドは非同期であるため、フェッチされたプロパティをコードで呼び出す前に、そのメソッドが完了していることを確認します。</span><span class="sxs-lookup"><span data-stu-id="d62fb-159">Because the `sync` method is asynchronous, ensure that it has completed before your code calls the properties that were fetched.</span></span>

<span data-ttu-id="d62fb-160">こうした手順は、コードで Office ドキュメントから情報を*読み取る*必要がある場合には必ず完了する必要があります。</span><span class="sxs-lookup"><span data-stu-id="d62fb-160">These steps must be completed whenever your code needs to *read* information from the Office document.</span></span>

1. <span data-ttu-id="d62fb-p112">`toggleProtection` 関数で、`TODO2` を次のコードに置き換えます。次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="d62fb-p112">In the `toggleProtection` function, replace `TODO2` with the following code. Note:</span></span>
   - <span data-ttu-id="d62fb-163">すべての Excel オブジェクトに `load` メソッドがあります。</span><span class="sxs-lookup"><span data-stu-id="d62fb-163">Every Excel object has a `load` method.</span></span> <span data-ttu-id="d62fb-164">読み取る必要のあるオブジェクトのプロパティは、コンマ区切りの名前の文字列としてパラメーターで指定します。</span><span class="sxs-lookup"><span data-stu-id="d62fb-164">You specify the properties of the object that you want to read in the parameter as a string of comma-delimited names.</span></span> <span data-ttu-id="d62fb-165">この場合、読み取る必要のあるプロパティは、`protection` プロパティのサブプロパティです。</span><span class="sxs-lookup"><span data-stu-id="d62fb-165">In this case, the property you need to read is a subproperty of the `protection` property.</span></span> <span data-ttu-id="d62fb-166">サブプロパティはその他のコードの場合とほとんど同じ方法で参照しますが、"." 記号の代わりにスラッシュ ('/') 記号を使用する点が異なります。</span><span class="sxs-lookup"><span data-stu-id="d62fb-166">You reference the subproperty almost exactly as you would anywhere else in your code, with the exception that you use a forward slash ('/') character instead of a "." character.</span></span>
   - <span data-ttu-id="d62fb-167">`sync` が完了してドキュメントからフェッチされた適切な値が `sheet.protection.protected` に割り当てられるまで、`sheet.protection.protected` を読み取る切り替えロジックが実行されないようにするために、そのロジックを `sync` が完了するまで実行されない `then` 関数に (この後の手順で) 移動します。</span><span class="sxs-lookup"><span data-stu-id="d62fb-167">To ensure that the toggle logic, which reads `sheet.protection.protected`, does not run until after the `sync` is complete and the `sheet.protection.protected` has been assigned the correct value that is fetched from the document, it will be moved (in the next step) into a `then` function that won't run until the `sync` has completed.</span></span> 

    ```javascript
    sheet.load('protection/protected');
    return context.sync()
        .then(
            function() {
                // TODO3: Move the queued toggle logic here.
            }
        )
        // TODO4: Move the final call of `context.sync` here and ensure that it
        //        does not run until the toggle logic has been queued.
    ``` 

2. <span data-ttu-id="d62fb-168">分岐していない同一のコード パスに 2 つの `return` ステートメントを含めることはできないため、`Excel.run` の最後にある最終行の `return context.sync();` を削除します。</span><span class="sxs-lookup"><span data-stu-id="d62fb-168">You can't have two `return` statements in the same unbranching code path, so delete the final line `return context.sync();` at the end of the `Excel.run`.</span></span> <span data-ttu-id="d62fb-169">この後の手順で、新しい最終の `context.sync` を追加します。</span><span class="sxs-lookup"><span data-stu-id="d62fb-169">You will add a new final `context.sync`, in a later step.</span></span>
3. <span data-ttu-id="d62fb-170">`toggleProtection` 関数内の `if ... else` 構造を切り取って、`TODO3` の代わりに貼り付けます。</span><span class="sxs-lookup"><span data-stu-id="d62fb-170">Cut the `if ... else` structure in the `toggleProtection` function and paste it in place of `TODO3`.</span></span>
4. <span data-ttu-id="d62fb-p115">`TODO4` を次のコードに置き換えます。次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="d62fb-p115">Replace `TODO4` with the following code. Note:</span></span>
   - <span data-ttu-id="d62fb-173">`sync` メソッドを `then` 関数に渡すことで、`sheet.protection.unprotect()` または `sheet.protection.protect()` のどちらかがキューに登録されるまで、そのメソッドが実行されないようにします。</span><span class="sxs-lookup"><span data-stu-id="d62fb-173">Passing the `sync` method to a `then` function ensures that it does not run until either `sheet.protection.unprotect()` or `sheet.protection.protect()` has been queued.</span></span>
   - <span data-ttu-id="d62fb-174">`then` メソッドは渡された関数を呼び出します。`sync` が 2 回呼び出されないように、`context.sync` の末尾の "()" は省略します。</span><span class="sxs-lookup"><span data-stu-id="d62fb-174">The `then` method invokes whatever function is passed to it, and you don't want `sync` to be invoked twice, so leave off the "()" from the end of `context.sync`.</span></span>

    ```javascript
    .then(context.sync);
    ```

   <span data-ttu-id="d62fb-175">作業が完了すると、関数の全体は次のようになります。</span><span class="sxs-lookup"><span data-stu-id="d62fb-175">When you are done, the entire function should look like the following:</span></span>

    ```javascript
    function toggleProtection(args) {
        Excel.run(function (context) {            
          const sheet = context.workbook.worksheets.getActiveWorksheet();          
          sheet.load('protection/protected');

          return context.sync()
              .then(
                  function() {
                    if (sheet.protection.protected) {
                        sheet.protection.unprotect();
                    } else {
                        sheet.protection.protect();
                    }
                  }
              )
              .then(context.sync);
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
        args.completed();
    }
    ```


## <a name="configure-the-script-loading-html-file"></a><span data-ttu-id="d62fb-176">スクリプト読み込み HTMl ファイルを構成する</span><span class="sxs-lookup"><span data-stu-id="d62fb-176">Configure the script-loading HTML file</span></span>

<span data-ttu-id="d62fb-177">/function-file/function-file.html ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="d62fb-177">Open the /function-file/function-file.html file.</span></span> <span data-ttu-id="d62fb-178">これは、ユーザーが **[Toggle Worksheet Protection]** ボタンをクリックしたときに呼び出される UI のない HTML ファイルです。</span><span class="sxs-lookup"><span data-stu-id="d62fb-178">This is a UI-less HTML file that is called when the user presses the **Toggle Worksheet Protection** button.</span></span> <span data-ttu-id="d62fb-179">ボタンがクリックされたときに実行する JavaScript メソッドを読み込むことを目的としています。</span><span class="sxs-lookup"><span data-stu-id="d62fb-179">Its purpose is to load the JavaScript method that should run when the button is pushed.</span></span> <span data-ttu-id="d62fb-180">このファイルには変更を加えません。</span><span class="sxs-lookup"><span data-stu-id="d62fb-180">You are not going to change this file.</span></span> <span data-ttu-id="d62fb-181">2 番目の `<script>` タグで functionfile.js が読み込まれる点に注目してください。</span><span class="sxs-lookup"><span data-stu-id="d62fb-181">Simply note that the second `<script>` tag loads the functionfile.js.</span></span>

   > [!NOTE]
   > <span data-ttu-id="d62fb-182">function-file.html ファイルと、そのファイルが読み込む function-file.js ファイルは、アドインの作業ウィンドウとは完全に別の IE プロセスで実行されます。</span><span class="sxs-lookup"><span data-stu-id="d62fb-182">The function-file.html file and the function-file.js file that it loads run in an entirely separate IE process from the add-in's task pane.</span></span> <span data-ttu-id="d62fb-183">function-file.js が app.js ファイルと同じ bundle.js ファイルからトランスパイルされていた場合、アドインでは bundle.js の 2 つのコピーを読み込むことが必要になり、バンドル化の意味がなくなります。</span><span class="sxs-lookup"><span data-stu-id="d62fb-183">If the function-file.js was transpiled into the same bundle.js file as the app.js file, then the add-in would have to load two copies of the bundle.js file, which defeats the purpose of bundling.</span></span> <span data-ttu-id="d62fb-184">さらに、function-file.js ファイルには IE で未サポートの JavaScript は含まれていません。</span><span class="sxs-lookup"><span data-stu-id="d62fb-184">In addition, the function-file.js file does not contain any JavaScript that is unsupported by IE.</span></span> <span data-ttu-id="d62fb-185">これら 2 つの理由から、このアドインでは function-file.js を一切トランスパイルしていません。</span><span class="sxs-lookup"><span data-stu-id="d62fb-185">For these two reasons, this add-in does not transpile the function-file.js at all.</span></span> 

## <a name="test-the-add-in"></a><span data-ttu-id="d62fb-186">アドインをテストする</span><span class="sxs-lookup"><span data-stu-id="d62fb-186">Test the add-in</span></span>

1. <span data-ttu-id="d62fb-187">Excel も含めて、すべての Office アプリケーションを閉じます。</span><span class="sxs-lookup"><span data-stu-id="d62fb-187">Close all Office applications, including Excel.</span></span> 
2. <span data-ttu-id="d62fb-188">キャッシュ フォルダーの内容を削除して、Office キャッシュを削除します。</span><span class="sxs-lookup"><span data-stu-id="d62fb-188">Delete the Office cache by deleting the contents of the cache folder.</span></span> <span data-ttu-id="d62fb-189">これは、ホストから古いバージョンのアドインを完全に削除するために必要です。</span><span class="sxs-lookup"><span data-stu-id="d62fb-189">This is necessary to completely clear the old version of the add-in from the host.</span></span> 
    - <span data-ttu-id="d62fb-190">Windows の場合: `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`。</span><span class="sxs-lookup"><span data-stu-id="d62fb-190">For Windows: `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`.</span></span>
    - <span data-ttu-id="d62fb-191">Mac の場合: `/Users/{your_name_on_the_device}/Library/Containers/com.Microsoft.OsfWebHost/Data/`。</span><span class="sxs-lookup"><span data-stu-id="d62fb-191">For Mac: `/Users/{your_name_on_the_device}/Library/Containers/com.Microsoft.OsfWebHost/Data/`.</span></span>
3. <span data-ttu-id="d62fb-192">何らかの理由で、サーバーが稼働中でない場合は、Git Bash ウィンドウ、または Node.JS 対応のシステム プロンプトで、プロジェクトの **Start** フォルダーに移動して、`npm start` コマンドを実行します。</span><span class="sxs-lookup"><span data-stu-id="d62fb-192">If for any reason, your server is not running, then in a Git Bash window, or Node.JS-enabled system prompt, navigate to the **Start** folder of the project and run the command `npm start`.</span></span> <span data-ttu-id="d62fb-193">変更した JavaScript ファイルはビルド済みの bundle.js に含まれていないため、プロジェクトをリビルドする必要はありません。</span><span class="sxs-lookup"><span data-stu-id="d62fb-193">You do not need to rebuild the project because the only JavaScript file you changed is not part of the built bundle.js.</span></span>
4. <span data-ttu-id="d62fb-194">新しいバージョンの変更済みマニフェスト ファイルを使用して、次のいずれかの方法でサイドローディング プロセスを繰り返します。</span><span class="sxs-lookup"><span data-stu-id="d62fb-194">Using the new version of the changed manifest file, repeat the sideloading process by using one of the following methods.</span></span> <span data-ttu-id="d62fb-195">*マニフェスト ファイルの以前のコピーを上書きする必要があります。*</span><span class="sxs-lookup"><span data-stu-id="d62fb-195">*You should overwrite the previous copy of the manifest file.*</span></span>
    - <span data-ttu-id="d62fb-196">Windows: [Windows で Office アドインをサイドロードする](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span><span class="sxs-lookup"><span data-stu-id="d62fb-196">Windows: [Sideload Office Add-ins on Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span></span>
    - <span data-ttu-id="d62fb-197">Excel Online:[Office Online で Office アドインをサイドロードする](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-online)</span><span class="sxs-lookup"><span data-stu-id="d62fb-197">Excel Online: [Sideload Office Add-ins in Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-online)</span></span>
    - <span data-ttu-id="d62fb-198">iPad および Mac: [iPad と Mac で Office アドインをサイドロードする](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span><span class="sxs-lookup"><span data-stu-id="d62fb-198">iPad and Mac: [Sideload Office Add-ins on iPad and Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span></span>
7. <span data-ttu-id="d62fb-199">Excel で任意のワークシートを開きます。</span><span class="sxs-lookup"><span data-stu-id="d62fb-199">Open any worksheet in Excel.</span></span>
8. <span data-ttu-id="d62fb-p121">**[ホーム]** リボンで、**[ワークシート保護の切り替え]** を選択します。次のスクリーンショットに示すように、リボンのほとんどのコントロールは、無効化 (淡色表示) されます。</span><span class="sxs-lookup"><span data-stu-id="d62fb-p121">On the **Home** ribbon, choose **Toggle Worksheet Protection**. Note that most of the controls on the ribbon are disabled (and visually grayed-out) as seen in screenshot below.</span></span> 
9. <span data-ttu-id="d62fb-202">セルの内容を変更する場合は、そのセルを選択します。</span><span class="sxs-lookup"><span data-stu-id="d62fb-202">Choose a cell as you would if you wanted to change its content.</span></span> <span data-ttu-id="d62fb-203">ワークシートが保護されているというエラーが表示されます。</span><span class="sxs-lookup"><span data-stu-id="d62fb-203">You get an error telling you that the worksheet is protected.</span></span>
10. <span data-ttu-id="d62fb-204">もう一度 **[Toggle Worksheet Protection]** を選択すると、コントロールが再有効化され、再びセルの値を変更できるようになります。</span><span class="sxs-lookup"><span data-stu-id="d62fb-204">Choose **Toggle Worksheet Protection** again, and the controls are reenabled, and you can change cell values again.</span></span>

    ![Excel チュートリアル: 保護がオンになっているリボン](../images/excel-tutorial-ribbon-with-protection-on.png)
