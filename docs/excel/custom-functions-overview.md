# <a name="create-custom-functions-in-excel-preview"></a><span data-ttu-id="ea082-101">Excel でのカスタム関数の作成 (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="ea082-101">Create custom functions in Excel (Preview)</span></span>

<span data-ttu-id="ea082-102">カスタム関数 (ユーザー定義関数 (UDF) と同様) を使用すると、開発者はアドインを使用して任意の JavaScript 関数を Excel に追加できます。</span><span class="sxs-lookup"><span data-stu-id="ea082-102">Custom functions (similar to user-defined functions, or UDFs), allow developers to add any JavaScript function to Excel using an add-in.</span></span> <span data-ttu-id="ea082-103">ユーザーは、Excel の他のネイティブ関数 (`=SUM()` など) と同様に、カスタム関数にアクセスできます。</span><span class="sxs-lookup"><span data-stu-id="ea082-103">Users can then access custom functions like any other native function in Excel (like =SUM()).</span></span> <span data-ttu-id="ea082-104">この記事では、Excel でカスタム関数を作成する方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="ea082-104">This article explains how to create custom functions in Excel.</span></span>

<span data-ttu-id="ea082-105">次の図は、エンドユーザーがカスタム関数をセルに挿入する方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="ea082-105">The following illustration shows you how an end user would insert a custom function into a cell.</span></span> <span data-ttu-id="ea082-106">1 組の数字に 42 を加える関数。</span><span class="sxs-lookup"><span data-stu-id="ea082-106">Here’s the code for a sample custom function that adds 42 to a pair of numbers.</span></span>

<img alt="custom functions" src="../images/custom-function.gif" width="579" height="383" />

<span data-ttu-id="ea082-107">同じカスタム関数のコードは次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="ea082-107">Here’s the code for the same custom function.</span></span>

```js
function ADD42(a, b) {
    return a + b + 42;
}
```

<span data-ttu-id="ea082-108">カスタム機能は、Windows、Mac、および Excel Online の開発者プレビューで利用できるようになりました。</span><span class="sxs-lookup"><span data-stu-id="ea082-108">Custom functions are now available in Developer Preview on Windows, Mac, and Excel Online.</span></span> <span data-ttu-id="ea082-109">以下の手順に従って試してみましょう。</span><span class="sxs-lookup"><span data-stu-id="ea082-109">Follow these steps to try them:</span></span>

1. <span data-ttu-id="ea082-110">Office（Windows では build 9325、Mac では 13.329）をインストールし、 [Office Insider](https://products.office.com/office-insider) プログラムに参加します。</span><span class="sxs-lookup"><span data-stu-id="ea082-110">Install Office (build 9325 on Windows or 13.329 on Mac) and join the [Office Insider](https://products.office.com/office-insider) program.</span></span> <span data-ttu-id="ea082-111">（最新のビルドを入手するだけでは不十分であることに注意してください。Insider プログラムに参加するまでは、どのビルドでも機能が無効になります）</span><span class="sxs-lookup"><span data-stu-id="ea082-111">(Note that it isn't enough just to get the latest build; the feature will be disabled on any build until you join the Insider program)</span></span>
2. <span data-ttu-id="ea082-112">[Yo Office](https://github.com/OfficeDev/generator-office) を使用して Excel カスタム関数のアドインを作成し、[ プロジェクトの README.md](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/README.md) の指示に従って Excel でアドインを起動し、コードを変更してデバッグします。</span><span class="sxs-lookup"><span data-stu-id="ea082-112">Create an Excel Custom Functions Add-in project using [Yo Office](https://github.com/OfficeDev/generator-office), and follow the instructions in the [project README.md](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/README.md) to start the add-in in Excel, make changes in the code, and debug.</span></span>
3. <span data-ttu-id="ea082-113">任意のセルに `=CONTOSO.ADD42(1,2)` を入力し、**Enter** を押してカスタム関数を実行します。</span><span class="sxs-lookup"><span data-stu-id="ea082-113">Type `=CONTOSO.ADD42(1,2)` into any cell, and press **Enter** to run the custom function.</span></span>

<span data-ttu-id="ea082-114">この記事の末尾にある **既知の問題** のセクションを参照してください。このセクションには、カスタム関数の現在の制約が記載されており、時間の経過に従って更新されます。</span><span class="sxs-lookup"><span data-stu-id="ea082-114">See the Known Issues section at the end of this article, which includes current limitations of custom functions and will be updated over time.</span></span>

## <a name="learn-the-basics"></a><span data-ttu-id="ea082-115">基本操作の説明</span><span class="sxs-lookup"><span data-stu-id="ea082-115">Learn the basics</span></span>

<span data-ttu-id="ea082-116">複製されたサンプル リポジトリには、次のファイルが表示されます。</span><span class="sxs-lookup"><span data-stu-id="ea082-116">In the cloned sample repo, you’ll see the following files:</span></span>

- <span data-ttu-id="ea082-117">**./src/customfunctions.js**: カスタム関数のコードが含まれています (`ADD42` 関数については、上出の単純なコード例をご覧ください)。</span><span class="sxs-lookup"><span data-stu-id="ea082-117">**customfunctions.js**, which contains the custom function code (see the simple code example above for the `ADD42` function).</span></span>
- <span data-ttu-id="ea082-118">**./config/customfunctions.json** カスタム関数について Excel に通知する登録 JSON が含まれています。</span><span class="sxs-lookup"><span data-stu-id="ea082-118">**customfunctions.json**, which contains the registration JSON that tells Excel about your custom function.</span></span> <span data-ttu-id="ea082-119">登録すると、ユーザーがセルに入力するときに表示される使用可能な関数のリストにカスタム関数が表示されます。</span><span class="sxs-lookup"><span data-stu-id="ea082-119">Registration makes your custom functions appear in the list of available functions displayed when users type in cells.</span></span>
- <span data-ttu-id="ea082-120">**./index.html** JS ファイルへの &lt;Script&gt; 参照を提供します。</span><span class="sxs-lookup"><span data-stu-id="ea082-120">**./index.html**, which provides a &lt;Script&gt; reference to the JS file.</span></span> <span data-ttu-id="ea082-121">このファイルでは、Excel の UI は表示されません。</span><span class="sxs-lookup"><span data-stu-id="ea082-121">This file does not display UI in Excel.</span></span>
- <span data-ttu-id="ea082-122">**./manifest.xml**: HTML、JavaScript、および JSON ファイルの場所を Excel に通知します。また、アドインと共にインストールされているすべてのカスタム関数の名前空間も指定します。</span><span class="sxs-lookup"><span data-stu-id="ea082-122">**customfunctions.xml**, which tells Excel the location of the HTML, JavaScript, and JSON files; and also specifies a namespace for all the custom functions that are installed with the add-in.</span></span>

### <a name="json-file-configcustomfunctionsjson"></a><span data-ttu-id="ea082-123">JSON ファイル (./config/customfunctions.json)</span><span class="sxs-lookup"><span data-stu-id="ea082-123">JSON file (./config/customfunctions.json)</span></span>

<span data-ttu-id="ea082-124">customfunctions.json の以下のコードは、同じ `ADD42` 関数のメタデータを指定します。</span><span class="sxs-lookup"><span data-stu-id="ea082-124">The following code in customfunctions.json specifies the metadata for the same `ADD42` function.</span></span>

> [!NOTE]
> <span data-ttu-id="ea082-125">この例で使用されていないオプションを含むJSONファイルの詳細な参照情報は、「[カスタム関数登録 JSON](custom-functions-json.md)」 にあります。。</span><span class="sxs-lookup"><span data-stu-id="ea082-125">Detailed reference information for the JSON file, including options not used in this example, is at [Custom Functions Registration JSON](custom-functions-json.md).</span></span>

<span data-ttu-id="ea082-126">この例では、以下のことに注意してください。</span><span class="sxs-lookup"><span data-stu-id="ea082-126">Note that for this example:</span></span>

- <span data-ttu-id="ea082-127">カスタム関数は1つしかないので、 `functions` ARRAY のメンバーも1つです。</span><span class="sxs-lookup"><span data-stu-id="ea082-127">There's only one custom function, so there's only one member of the `functions` array.</span></span>
- <span data-ttu-id="ea082-128">プロパティは関数名を定義します。`name`</span><span class="sxs-lookup"><span data-stu-id="ea082-128">The `name` property defines the function name.</span></span> <span data-ttu-id="ea082-129">前に示したアニメーションGIFのように、名前空間（`CONTOSO`）は、Excel オートコンプリート メニューの関数名の前に付加されます。</span><span class="sxs-lookup"><span data-stu-id="ea082-129">As you see in the animated gif shown previously, a namespace (`CONTOSO`) is prepended to the function name in the Excel autocomplete menu.</span></span> <span data-ttu-id="ea082-130">このプレフィックスは、後述するアドインマニフェストで定義されます。</span><span class="sxs-lookup"><span data-stu-id="ea082-130">This prefix is defined in the add-in manifest, described below.</span></span> <span data-ttu-id="ea082-131">プレフィックスと関数名はピリオドで区切られ、慣例では接頭辞と関数名は大文字です。</span><span class="sxs-lookup"><span data-stu-id="ea082-131">The prefix and the function name are separated using a period, and by convention prefixes and function names are uppercase.</span></span> <span data-ttu-id="ea082-132">カスタム関数を使用するには、ユーザーが名前空間に続けて関数の名前（`ADD42` ）をセルに入力します。この場合、 `=CONTOSO.ADD42` です。</span><span class="sxs-lookup"><span data-stu-id="ea082-132">To use your custom function, a user types the namespace followed by the function's name (`ADD42`) into a cell, in this case `=CONTOSO.ADD42`.</span></span> <span data-ttu-id="ea082-133">プレフィックスは、所属する会社やアドインの識別子として使用することが想定されています。</span><span class="sxs-lookup"><span data-stu-id="ea082-133">The prefix is intended to be used as an identifier for your add-in.</span></span> 
- <span data-ttu-id="ea082-134">Excel のオートコンプリート メニュー `description` 表示されます。</span><span class="sxs-lookup"><span data-stu-id="ea082-134">`description`: The description appears in the autocomplete menu in Excel.</span></span>
- <span data-ttu-id="ea082-135">ユーザーが関数のヘルプを要求すると、Excel は作業ウィンドウを開き、`helpUrl` に指定された URL にある Web ページを表示します。</span><span class="sxs-lookup"><span data-stu-id="ea082-135">`helpUrl`: When the user requests help for a function, Excel opens a task pane and displays the web page found at this URL.</span></span>
- <span data-ttu-id="ea082-136">`result` プロパティは、関数が Excel に返す情報の種類を指定します。</span><span class="sxs-lookup"><span data-stu-id="ea082-136">`result`: Defines the type of information returned by the function to Excel.</span></span> <span data-ttu-id="ea082-137">子のプロパティは `"string"`、 `"number"`、または `"boolean"` ができます。。`type`</span><span class="sxs-lookup"><span data-stu-id="ea082-137">The `type` child property can `"string"`, `"number"`, or `"boolean"`.</span></span> <span data-ttu-id="ea082-138">プロパティは `scalar` または `matrix` （指定された`type` の値の2次元配列）とすることができます。`dimensionality`</span><span class="sxs-lookup"><span data-stu-id="ea082-138">The `dimensionality` property can be `scalar` or `matrix` (a two-dimensional array of values of the specified `type`.)</span></span>
- <span data-ttu-id="ea082-139">配列は、 関数に渡される各パラメーターのデータの種類を *順番に* 指定します。`parameters`</span><span class="sxs-lookup"><span data-stu-id="ea082-139">The `parameters` array specifies, *in order*, the type of data in each parameter that is passed to the function.</span></span> <span data-ttu-id="ea082-140">と `description` 子のプロパティは Excel intellisense で使用されます。`name`</span><span class="sxs-lookup"><span data-stu-id="ea082-140">The `name` and `description` child properties are used in the Excel intellisense.</span></span> <span data-ttu-id="ea082-141">と `dimensionality` 子のプロパティは上記で説明した `result` プロパティの子プロパティと同じです。`type`</span><span class="sxs-lookup"><span data-stu-id="ea082-141">The `type` and `dimensionality` child properties are identical to the child properties of the `result` property described above.</span></span>
- <span data-ttu-id="ea082-142">プロパティを使用すると、Excel がいつどのようにして関数を実行するかについてのいくつかの側面をカスタマイズできます。`options`</span><span class="sxs-lookup"><span data-stu-id="ea082-142">The `options` property enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="ea082-143">これらのオプションについての詳細がこの記事の後半にあります。</span><span class="sxs-lookup"><span data-stu-id="ea082-143">There is more information about these options later in this article.</span></span>

```js
    {
        "$schema": "https://developer.microsoft.com/json-schemas/office-js/custom-functions.schema.json",
        "functions": [
            {
                "name": "ADD42", 
                "description":  "adds 42 to the input numbers",
                "helpUrl": "http://dev.office.com",
                "result": {
                    "type": "number",
                    "dimensionality": "scalar"
                },
                "parameters": [
                    {
                        "name": "number 1",
                        "description": "the first number to be added",
                        "type": "number",
                        "dimensionality": "scalar"
                    },
                    {
                        "name": "number 2",
                        "description": "the second number to be added",
                        "type": "number",
                        "dimensionality": "scalar"
                    }
                ],
                "options": {
                    "sync": true
                }
            }
        ]
    }
```

> [!NOTE]
> <span data-ttu-id="ea082-144">カスタム関数は、ユーザーが最初にアドインを実行したときに登録されます。</span><span class="sxs-lookup"><span data-stu-id="ea082-144">The custom functions are registered when a user runs the add-in for the first time.</span></span> <span data-ttu-id="ea082-145">その後、同じユーザーに対して、すべてのブック（アドインが最初に実行されたものだけでなく）で関数を使用できます。</span><span class="sxs-lookup"><span data-stu-id="ea082-145">After that, they are available, for that same user, in all workbooks (not only the one where the add-in ran initially.)</span></span>

<span data-ttu-id="ea082-146">JSON ファイルのサーバー設定では、カスタム関数が Excel Online で正しく作動するために [CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS) が有効になっていなければなりません。</span><span class="sxs-lookup"><span data-stu-id="ea082-146">Your server settings for the JSON file must have [CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS) enabled in order for custom functions to work correctly in Excel Online.</span></span>


### <a name="manifest-file-manifestxml"></a><span data-ttu-id="ea082-147">マニフェスト ファイル (./manifest.xml)</span><span class="sxs-lookup"><span data-stu-id="ea082-147">Manifest file (manifest.xml)</span></span>


<span data-ttu-id="ea082-148">以下は、作成した関数を Excel が実行できるようにアドインのマニフェストに組み込んだ `<ExtensionPoint>` および `<Resources>` マークアップの例です。</span><span class="sxs-lookup"><span data-stu-id="ea082-148">The following is an example of the `<ExtensionPoint>` and `<Resources>` markup that you include in the add-in's manifest to enable Excel to run your functions.</span></span> <span data-ttu-id="ea082-149">このマークアップについて、次の点にご注意ください。</span><span class="sxs-lookup"><span data-stu-id="ea082-149">Note the following facts about this markup:</span></span>

- <span data-ttu-id="ea082-150">要素とそれに対応するリソース ID は、関数で JavaScript ファイルの場所を指定します。`<Script>`</span><span class="sxs-lookup"><span data-stu-id="ea082-150">The `<Script>` element and its corresponding resource ID specifies the location of the JavaScript file with your functions.</span></span>
- <span data-ttu-id="ea082-151">要素とそれに対応するリソース ID は、アドインの HTML ページの場所を指定します。`<Page>`</span><span class="sxs-lookup"><span data-stu-id="ea082-151">The `<Page>` element and its corresponding resource ID specifies the location of the HTML page of your add-in.</span></span> <span data-ttu-id="ea082-152">HTML ページには、JavaScript ファイル（customfunctions.js）を読み込む `<Script>` タグが含まれています。</span><span class="sxs-lookup"><span data-stu-id="ea082-152">The HTML page includes a `<Script>` tag that loads the JavaScript file (customfunctions.js).</span></span> <span data-ttu-id="ea082-153">HTML ページは非表示のページであり、UI に表示されることはありません。</span><span class="sxs-lookup"><span data-stu-id="ea082-153">The HTML page is a hidden page and is never displayed in the UI.</span></span>
- <span data-ttu-id="ea082-154">要素とそれに対応するリソース ID は、JSON ファイルの場所を指定します。`<Metadata>`</span><span class="sxs-lookup"><span data-stu-id="ea082-154">The `<Metadata>` element and its corresponding resource ID specifies the location of the JSON file.</span></span>
- <span data-ttu-id="ea082-155">要素および対応するリソース ID は、アドインのすべてのカスタム関数のプレフィックスを指定します。`<Namespace>`</span><span class="sxs-lookup"><span data-stu-id="ea082-155">A `<Namespace>` element and its corresponding resource ID specifies the prefix for all custom functions in the add-in.</span></span>


```xml
<VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1\_0">
    <Hosts>
        <Host xsi:type="Workbook">
            <AllFormFactors>
                <ExtensionPoint xsi:type="CustomFunctions">
                    <Script>
                        <SourceLocation resid="residjs" />
                    </Script>
                    <Page>
                        <SourceLocation resid="residhtml"/>
                    </Page>
                    <Metadata>
                        <SourceLocation resid="residjson" />
                    </Metadata>
                    <Namespace resid="residNS" />
                </ExtensionPoint>
            </AllFormFactors>
        </Host>
    </Hosts>
    <Resources>
        <bt:Urls>
            <bt:Url id="residjson" DefaultValue="http://127.0.0.1:8080/customfunctions.json" />
            <bt:Url id="residjs" DefaultValue="http://127.0.0.1:8080/customfunctions.js" />
            <bt:Url id="residhtml" DefaultValue="http://127.0.0.1:8080/customfunctions.html" />
        </bt:Urls>
        <bt:ShortStrings>
            <bt:String id="residNS" DefaultValue="CONTOSO" />
        </bt:ShortStrings>
    </Resources>
</VersionOverrides>

```

## <a name="initializing-custom-functions"></a><span data-ttu-id="ea082-156">カスタム関数の初期化</span><span class="sxs-lookup"><span data-stu-id="ea082-156">Initializing custom functions</span></span>

<span data-ttu-id="ea082-157">コードは、使用する前にカスタム関数の機能を初期化する必要があります。</span><span class="sxs-lookup"><span data-stu-id="ea082-157">Your code must initialize the custom functions feature before using it.</span></span> <span data-ttu-id="ea082-158">初期化は、HTML ファイル （customfunctions.html）の &lt;Script&gt; タグ、または JavaScript ファイル（customfuntions.js）のトップで実行できます。</span><span class="sxs-lookup"><span data-stu-id="ea082-158">You can do this either in a &lt;Script&gt; tag in the HTML file (customfunctions.html) or at the top of the JavaScript file (customfunctions.js).</span></span> <span data-ttu-id="ea082-159">カスタム関数のプレビュー中に、初期化のための 2 つの構文を選択できます。</span><span class="sxs-lookup"><span data-stu-id="ea082-159">During the preview of custom functions, you have your choice of two syntaxes for intializing.</span></span> <span data-ttu-id="ea082-160">リポジトリ内の HTML ファイルは、次の構文を使用します。</span><span class="sxs-lookup"><span data-stu-id="ea082-160">The HTML file in the repo uses the following syntax:</span></span>

```js
Office.initialize = function (reason) {
    return Excel.CustomFunctions.initialize();
};
```

<span data-ttu-id="ea082-161">次の構文も使用できます。</span><span class="sxs-lookup"><span data-stu-id="ea082-161">You can also use the following syntax:</span></span>

```js
Office.Preview.StartCustomFunctions();
```

## <a name="handling-errors"></a><span data-ttu-id="ea082-162">エラーの処理</span><span class="sxs-lookup"><span data-stu-id="ea082-162">Handling errors</span></span>
<span data-ttu-id="ea082-163">カスタム関数のエラー処理は、[一般的な Excel JavaScript API のエラー処理](./excel-add-ins-error-handling.md) と同じです。</span><span class="sxs-lookup"><span data-stu-id="ea082-163">Error handling for custom functions is the same as [error handling for the Excel JavaScript API at large](./excel-add-ins-error-handling.md).</span></span> <span data-ttu-id="ea082-164">一般的に、エラー処理には `.catch` を使用します。</span><span class="sxs-lookup"><span data-stu-id="ea082-164">Generally, you will use `.catch` to handle errors.</span></span> <span data-ttu-id="ea082-165">次のコードは、`.catch` の例を示しています。</span><span class="sxs-lookup"><span data-stu-id="ea082-165">The code below gives an example of `.catch`.</span></span> 

```js
function getComment(x) {
    var url = "https://jsonplaceholder.typicode.com/comments/" + x; //this delivers a section of lorem ipsum from the jsonplaceholder API
    return fetch(url)
        .then(function (data) {
            return data.json();
        })
        .then((json) => {
            return json.body;
        })
        .catch(function (error) {
            throw error;
        })
}
```

## <a name="synchronous-and-asynchronous-functions"></a><span data-ttu-id="ea082-166">同期関数と非同期関数</span><span class="sxs-lookup"><span data-stu-id="ea082-166">Synchronous and asynchronous processing</span></span>

<span data-ttu-id="ea082-167">上記の `ADD42` 関数は Excel （JSON ファイルのオプション `"sync": true` を使用して指定したもの ）と同期しています。</span><span class="sxs-lookup"><span data-stu-id="ea082-167">The function `ADD42` above is synchronous with respect to Excel (designated by setting the option `"sync": true` in the JSON file).</span></span> <span data-ttu-id="ea082-168">同期関数は、Excel と同じプロセスで実行され、マルチスレッド計算中に並行して実行されるため、高速なパフォーマンスを提供します。</span><span class="sxs-lookup"><span data-stu-id="ea082-168">Synchronous functions offer fast performance because they run in the same process as Excel and they run in parallel during multithreaded calculation.</span></span>   

<span data-ttu-id="ea082-169">一方、カスタム関数が Web からデータを取得する場合は、Excel と非同期でなければなりません。</span><span class="sxs-lookup"><span data-stu-id="ea082-169">On the other hand, if your custom function retrieves data from the web, it must be asynchronous with respect to Excel.</span></span> <span data-ttu-id="ea082-170">非同期関数は以下を実行する必要があります。</span><span class="sxs-lookup"><span data-stu-id="ea082-170">Asynchronous functions must:</span></span>

1. <span data-ttu-id="ea082-171">JavaScript Promise を Excel に返します。</span><span class="sxs-lookup"><span data-stu-id="ea082-171">Return a JavaScript Promise to Excel.</span></span>
3. <span data-ttu-id="ea082-172">コールバック関数を使用して Promise を最終値で解決します。</span><span class="sxs-lookup"><span data-stu-id="ea082-172">Resolve the Promise with the final value using the callback function.</span></span>

<span data-ttu-id="ea082-173">次のコードは、温度計の温度を取得する非同期カスタム関数の例を示しています。</span><span class="sxs-lookup"><span data-stu-id="ea082-173">The following code shows an example of a custom function that retrieves the temperature of a thermometer.</span></span> <span data-ttu-id="ea082-174">は、XHR を使用して温度 Web サービスを呼び出す、ここでは指定されていない仮想関数であることにご注意ください。`sendWebRequest`</span><span class="sxs-lookup"><span data-stu-id="ea082-174">Note that `sendWebRequest` is a hypothetical function, not specified here, that uses XHR to call a temperature web service.</span></span>

```js
function getTemperature(thermometerID){
    return new OfficeExtension.Promise(function(setResult){
        sendWebRequest(thermometerID, function(data){
            setResult(data.temperature);
        });
    });
}
```

<span data-ttu-id="ea082-175">非同期関数は、 Excelが最終結果を待つ間、セルに `GETTING_DATA` 一時的エラーを表示します。</span><span class="sxs-lookup"><span data-stu-id="ea082-175">Asynchronous functions display a `GETTING_DATA` temporary error in the cell while Excel waits for the final result.</span></span> <span data-ttu-id="ea082-176">ユーザーは、結果を待つ間、スプレッドシートの他の部分と通常通りやりとりすることができます。</span><span class="sxs-lookup"><span data-stu-id="ea082-176">Users can interact normally with the rest of the spreadsheet while they wait for the result.</span></span>

> [!NOTE]
> <span data-ttu-id="ea082-177">カスタム関数は既定では非同期です。</span><span class="sxs-lookup"><span data-stu-id="ea082-177">Custom functions are asynchronous by default.</span></span> <span data-ttu-id="ea082-178">同期として関数を指定するには、登録 JSON ファイル内のカスタム関数の `options` プロパティでオプション `"sync": true` を設定してください。</span><span class="sxs-lookup"><span data-stu-id="ea082-178">To designate functions as synchronous set the option `"sync": true` in the `options` property for the custom function in the registration JSON file.</span></span>

## <a name="streamed-functions"></a><span data-ttu-id="ea082-179">ストリーム関数</span><span class="sxs-lookup"><span data-stu-id="ea082-179">Streamed functions</span></span>

<span data-ttu-id="ea082-180">非同期関数をストリーミングできます。</span><span class="sxs-lookup"><span data-stu-id="ea082-180">An asynchronous function can be streamed.</span></span> <span data-ttu-id="ea082-181">カスタムのストリーム関数を使用すると、Excel やユーザーが再計算を要求するのを待たずに、時間の経過に従ってセルに繰り返しデータを出力できます。</span><span class="sxs-lookup"><span data-stu-id="ea082-181">Streamed custom functions let you output data to cells repeatedly over time, without waiting for Excel or users to request recalculations.</span></span> <span data-ttu-id="ea082-182">次の例は、1 秒おきに結果に数値を追加するカスタム関数です。</span><span class="sxs-lookup"><span data-stu-id="ea082-182">The following example is a custom function that adds a number to the result every second.</span></span> <span data-ttu-id="ea082-183">このコードについては、次の点にご注意ください。</span><span class="sxs-lookup"><span data-stu-id="ea082-183">Note the following about this code:</span></span>

- <span data-ttu-id="ea082-184">Excel は、`setResult`コールバックを使用して自動的に新しい値を表示します。</span><span class="sxs-lookup"><span data-stu-id="ea082-184">Excel displays each new value automatically using the `setResult` callback.</span></span>
- <span data-ttu-id="ea082-185">|||UNTRANSLATED_CONTENT_START|||The final parameter, `handler`, is never specified in your registration code, and it does not display in the autocomplete menu to Excel users when they enter the function.|||UNTRANSLATED_CONTENT_END|||</span><span class="sxs-lookup"><span data-stu-id="ea082-185">For streamed functions, the final parameter, `handler`, is never specified in your registration code, and it does not display in the autocomplete menu to Excel users when they enter the function.</span></span> <span data-ttu-id="ea082-186">これは、関数のデータを Excel に渡してセルの値を更新するために使用される `setResult` コールバック関数を含むオブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="ea082-186">It’s an object that contains a `setResult` callback function that’s used to pass data from the function to Excel to update the value of a cell.</span></span>
- <span data-ttu-id="ea082-187">Excel が `handler` オブジェクト内の `setResult` 関数を渡すには、関数登録の際に、登録 JSON ファイル内のカスタム関数の `options` プロパティでオプション `"stream": true` を設定して、ストリーミングのサポートを宣言する必要があります。</span><span class="sxs-lookup"><span data-stu-id="ea082-187">In order for Excel to pass the `setResult` function in the `handler` object, you must declare support for streaming during your function registration by setting the option `"stream": true` in the `options` property for the custom function in the registration JSON file.</span></span>

```js
function incrementValue(increment, handler){
    var result = 0;
    setInterval(function(){
         result += increment;
         handler.setResult(result);
    }, 1000);
}
```

## <a name="cancellation"></a><span data-ttu-id="ea082-188">キャンセル</span><span class="sxs-lookup"><span data-stu-id="ea082-188">Cancellation</span></span>

<span data-ttu-id="ea082-189">ストリーム関数と非同期関数をキャンセルできます。</span><span class="sxs-lookup"><span data-stu-id="ea082-189">You can cancel streamed functions and asynchronous functions.</span></span> <span data-ttu-id="ea082-190">関数呼び出しのキャンセルは、帯域幅の使用量、作業メモリ、および CPU の負荷を減らすために重要です。</span><span class="sxs-lookup"><span data-stu-id="ea082-190">Canceling your function calls is important to reduce their bandwith consumption, working memory, and CPU load.</span></span> <span data-ttu-id="ea082-191">Excel では、次のような状況で関数の呼び出しをキャンセルします。</span><span class="sxs-lookup"><span data-stu-id="ea082-191">Excel cancels function calls in the following situations:</span></span>

- <span data-ttu-id="ea082-192">ユーザーが関数を参照するセルを編集または削除する。</span><span class="sxs-lookup"><span data-stu-id="ea082-192">The user edits or deletes a cell that references the function.</span></span>
- <span data-ttu-id="ea082-193">関数の引数 (入力) の 1 つが変更される。</span><span class="sxs-lookup"><span data-stu-id="ea082-193">One of the arguments (inputs) for the function changes.</span></span> <span data-ttu-id="ea082-194">この場合、キャンセルに加えて新しい関数の呼び出しがトリガーされます。</span><span class="sxs-lookup"><span data-stu-id="ea082-194">In this case, a new function call is triggered in addition to the cancelation.</span></span>
- <span data-ttu-id="ea082-p125">ユーザーは手動で再計算をトリガーします。上記の場合と同様に、キャンセルに加えて新しい関数の呼び出しがトリガーされます。</span><span class="sxs-lookup"><span data-stu-id="ea082-p125">The user triggers recalculation manually. As with the above case, a new function call is triggered in addition to the cancelation.</span></span>

<span data-ttu-id="ea082-197">すべてのストリーミング関数に対してキャンセル ハンドラを実装することが *必須* です。</span><span class="sxs-lookup"><span data-stu-id="ea082-197">You *must* implement a cancellation handler for every streaming function.</span></span> <span data-ttu-id="ea082-198">非同期の非ストリーミング関数は、キャンセル可能にもキャンセル不可にもでき、ご自分で決定できます。</span><span class="sxs-lookup"><span data-stu-id="ea082-198">Asynchronous, non-streaming functions may or may not be cancelable; it's up to you.</span></span> <span data-ttu-id="ea082-199">同期機能はキャンセルすることはできません。</span><span class="sxs-lookup"><span data-stu-id="ea082-199">Synchronous functions cannot be canceled.</span></span>

<span data-ttu-id="ea082-200">関数をキャンセル可能にするには、登録 JSON ファイル内のカスタム関数の `options` プロパティでオプション `"cancelable": true` を設定してください。</span><span class="sxs-lookup"><span data-stu-id="ea082-200">To make a function cancelable, set the option `"cancelable": true` in the `options` property for the custom function in the registration JSON file.</span></span>

<span data-ttu-id="ea082-201">次のコードでは、前述の例にキャンセルを実装しています。</span><span class="sxs-lookup"><span data-stu-id="ea082-201">The following code shows the previous example with cancellation implemented.</span></span> <span data-ttu-id="ea082-202">このコードでは、`handler` オブジェクトに `onCanceled` 関数が含まれており、キャンセル可能な各カスタム関数ごとにこの関数を定義する必要があります。</span><span class="sxs-lookup"><span data-stu-id="ea082-202">In the code, the `handler` object contains an `onCanceled` function which should be defined for each custom function.</span></span>

```js
function incrementValue(increment, handler){ 
    var result = 0;
    var timer = setInterval(function(){
         result += increment;
         handler.setResult(result);
    }, 1000);

    handler.onCanceled = function(){
        clearInterval(timer);
    }
}
```

## <a name="saving-and-sharing-state"></a><span data-ttu-id="ea082-203">状態の保存と共有</span><span class="sxs-lookup"><span data-stu-id="ea082-203">Saving and sharing state</span></span>

<span data-ttu-id="ea082-204">非同期カスタム関数では、JavaScript のグローバル変数にデータを格納できます。</span><span class="sxs-lookup"><span data-stu-id="ea082-204">Custom functions can save data in global JavaScript variables.</span></span> <span data-ttu-id="ea082-205">後続の呼び出しでは、カスタム関数はこれらの変数に格納されている値を使用できます。</span><span class="sxs-lookup"><span data-stu-id="ea082-205">In subsequent calls, your custom function may use the values saved in these variables.</span></span> <span data-ttu-id="ea082-206">保存された状態は、関数のすべてのインスタンスが状態を共有できるため、ユーザーが複数のセルに同じカスタム関数を追加する場合に便利です。</span><span class="sxs-lookup"><span data-stu-id="ea082-206">Saved state is useful when users add the same custom function to more than one cell, because all the instances of the function can share the state.</span></span> <span data-ttu-id="ea082-207">たとえば、同じ Web リソースへの追加呼び出しを避けるために、呼び出しから返されたデータを Web リソースに保存することができます。</span><span class="sxs-lookup"><span data-stu-id="ea082-207">For example, you may save the data returned from a call to a web resource to avoid making additional calls to the same web resource.</span></span>

<span data-ttu-id="ea082-208">次のコードは、 状態をグローバルで格納する前述の温度ストリーミング関数の実装を示しています。</span><span class="sxs-lookup"><span data-stu-id="ea082-208">The following code shows an implementation of the previous temperature-streaming function that saves state using the  variable.</span></span> <span data-ttu-id="ea082-209">このコードについては、次の点にご注意ください。</span><span class="sxs-lookup"><span data-stu-id="ea082-209">Note the following about this code:</span></span>

- <span data-ttu-id="ea082-210">`refreshTemperature` は、1 秒おきに特定の温度計の温度を読み取るストリーム関数です。</span><span class="sxs-lookup"><span data-stu-id="ea082-210">`refreshTemperature` is a streamed function that reads the temperature of a particular thermometer every second.</span></span> <span data-ttu-id="ea082-211">新しい温度は `savedTemperatures` 変数に保存されますが、セルの値を直接更新することはありません。</span><span class="sxs-lookup"><span data-stu-id="ea082-211">New temperatures are saved in the `savedTemperatures` variable, but does not directly update the cell value.</span></span> <span data-ttu-id="ea082-212">ワークシート・セルから直接呼び出されません。\*したがって、JSON ファイルには登録されません \*</span><span class="sxs-lookup"><span data-stu-id="ea082-212">It should not be directly called from a worksheet cell, *so it is not registered in the JSON file*.</span></span>
- <span data-ttu-id="ea082-213">`streamTemperature` 1 秒おきにセルに表示される温度値を更新します。また、 `savedTemperatures` 変数をデータソースとして使用します。</span><span class="sxs-lookup"><span data-stu-id="ea082-213">`streamTemperature` updates the temperature values displayed in the cell every second and it uses `savedTemperatures` variable as its data source.</span></span> <span data-ttu-id="ea082-214">JSON ファイルに登録し、すべて大文字で `STREAMTEMPERATURE` という名前をつける必要があります。</span><span class="sxs-lookup"><span data-stu-id="ea082-214">It must be registered in the JSON file, and named with all upper-case letters, `STREAMTEMPERATURE`.</span></span>
- <span data-ttu-id="ea082-215">ユーザーは、Excel UI の複数のセルから `streamTemperature` を呼び出すことができます。</span><span class="sxs-lookup"><span data-stu-id="ea082-215">Users may call `streamTemperature` from several cells in the Excel UI.</span></span> <span data-ttu-id="ea082-216">呼び出すたびに、同じ `savedTemperatures` 変数からデータを読み取ります。</span><span class="sxs-lookup"><span data-stu-id="ea082-216">Each call reads data from the same `savedTemperatures` variable.</span></span>

```js
var savedTemperatures;

function streamTemperature(thermometerID, handler){ 
     if(!savedTemperatures[thermometerID]){
         refreshTemperatures(thermometerID); // starts fetching temperatures if the thermometer hasn't been read yet
     }

     function getNextTemperature(){
         handler.setResult(savedTemperatures[thermometerID]); // setResult sends the saved temperature value to Excel.
         setTimeout(getNextTemperature, 1000); // Wait 1 second before updating Excel again.
     }
     getNextTemperature();
}

function refreshTemperature(thermometerID){
     sendWebRequest(thermometerID, function(data){
         savedTemperatures[thermometerID] = data.temperature;
     });
     setTimeout(function(){
         refreshTemperature(thermometerID);
     }, 1000); // Wait 1 second before reading the thermometer again, and then update the saved temperature of thermometerID.
}
```

> [!NOTE]
> <span data-ttu-id="ea082-217">同期関数（JSON ファイル内のオプション `"sync": true` で指定されたもの）は、Excel がマルチスレッド計算中にそれらを並行して行うため、状態を共有できません。</span><span class="sxs-lookup"><span data-stu-id="ea082-217">Synchronous functions (designated by setting the option `"sync": true` in the JSON file) cannot share state because Excel parallelizes them during multithreaded calculation.</span></span> <span data-ttu-id="ea082-218">アドインの同期関数が各セッションで同じ JavaScript コンテキストを共有するため、非同期関数のみが状態を共有できます。</span><span class="sxs-lookup"><span data-stu-id="ea082-218">Only asynchronous functions may share state because an add-in's synchronous functions share the same JavaScript context in each session.</span></span>

## <a name="working-with-ranges-of-data"></a><span data-ttu-id="ea082-219">データの範囲を使用する</span><span class="sxs-lookup"><span data-stu-id="ea082-219">Working with ranges of data</span></span>

<span data-ttu-id="ea082-220">カスタム関数は、データ範囲をパラメーターとして受け取ったり、カスタム関数からデータ範囲を返したりすることができます。</span><span class="sxs-lookup"><span data-stu-id="ea082-220">Your custom function can take a range of data as a parameter, or you can return a range of data from a custom function.</span></span>

<span data-ttu-id="ea082-221">たとえば、関数が Excel に格納されている数値の範囲から 2 番目に高い値を返すとします。</span><span class="sxs-lookup"><span data-stu-id="ea082-221">For example, suppose that your function returns the second highest temperature from a range of temperature values stored in Excel.</span></span> <span data-ttu-id="ea082-222">次の関数は、パラメーター `values` を受け取ります。これは `Excel.CustomFunctionDimensionality.matrix` パラメーター型です。</span><span class="sxs-lookup"><span data-stu-id="ea082-222">The following function takes the parameter `values`, which is an `Excel.CustomFunctionDimensionality.matrix` parameter type.</span></span> <span data-ttu-id="ea082-223">この関数の登録 JSON では、パラメータの `type` プロパティを `matrix` に設定するよう注意してください。</span><span class="sxs-lookup"><span data-stu-id="ea082-223">Note that in the registration JSON for this function, you would set the parameter's `type` property to `matrix`.</span></span>

```js
function secondHighest(values){ 
     var highest = values[0][0], secondHighest = values[0][0];
     for(var i = 0; i < values.length; i++){
         for(var j = 1; j < values[i].length; j++){
             if(values[i][j] >= highest){
                 secondHighest = highest;
                 highest = values[i][j];
             }
             else if(values[i][j] >= secondHighest){
                 secondHighest = values[i][j];
             }
         }
     }
     return secondHighest;
 }
```

<span data-ttu-id="ea082-224">ご覧のとおり、範囲は JavaScript で行配列の配列（2次元配列など）として処理されます。</span><span class="sxs-lookup"><span data-stu-id="ea082-224">As you can see, ranges are handled in JavaScript as arrays of row arrays (like a 2-dimensional array).</span></span>

## <a name="known-issues"></a><span data-ttu-id="ea082-225">既知の問題</span><span class="sxs-lookup"><span data-stu-id="ea082-225">Known issues</span></span>

- <span data-ttu-id="ea082-226">ヘルプの URL とパラメーターの説明は、Excel ではまだ使用されていません。</span><span class="sxs-lookup"><span data-stu-id="ea082-226">Help URLs and parameter descriptions are not yet used by Excel.</span></span>
- <span data-ttu-id="ea082-227">カスタム機能は現在、モバイル クライアント用の Excel では使用できません。</span><span class="sxs-lookup"><span data-stu-id="ea082-227">Custom functions are not currently available on Excel for mobile clients.</span></span>
- <span data-ttu-id="ea082-228">現在、アドインは、非同期関数カスタム関数を実行するために非表示ブラウザ プロセスを利用しています。</span><span class="sxs-lookup"><span data-stu-id="ea082-228">Currently, add-ins rely on a hidden browser process to run custom functions.</span></span> <span data-ttu-id="ea082-229">カスタム関数をより高速にし、使用メモリを少なくするために、今後 JavaScript は一部のプラットフォームでは直接実行されるようになります。</span><span class="sxs-lookup"><span data-stu-id="ea082-229">In the future, JavaScript will run directly on some platforms to ensure custom functions are faster and use less memory.</span></span> <span data-ttu-id="ea082-230">さらに、マニフェストの `<Page>` 要素によって参照される HTML ページは、Excel が JavaScript を直接実行するようになるため、ほとんどのプラットフォームで不要になります。</span><span class="sxs-lookup"><span data-stu-id="ea082-230">Additionally, the HTML page referenced by the `<Page>`Page element in the manifest won’t be needed for most platforms because Excel will run the JavaScript directly.</span></span> <span data-ttu-id="ea082-231">この変更に備えるため、カスタム関数が Web ページ DOM を使用しないことを徹底してください。</span><span class="sxs-lookup"><span data-stu-id="ea082-231">To prepare for this change, ensure your custom functions do not use the webpage DOM.</span></span> <span data-ttu-id="ea082-232">Web にアクセスするためにサポートされているホスト API は、GET または POST を使用する [WebSocket](https://developer.mozilla.org/docs/Web/API/WebSockets_API) および [XHR](https://developer.mozilla.org/docs/Web/API/XMLHttpRequest) になります。</span><span class="sxs-lookup"><span data-stu-id="ea082-232">The supported host APIs for accessing the web will be [WebSocket](https://developer.mozilla.org/docs/Web/API/WebSockets_API) and [XHR](https://developer.mozilla.org/docs/Web/API/XMLHttpRequest) using GET or POST.</span></span>
- <span data-ttu-id="ea082-233">揮発性関数（スプレッドシート内の無関係なデータが変更されたときに自動的に再計算する関数）はまだサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="ea082-233">Volatile functions (those which recalculate automatically whenever unrelated data changes in the spreadsheet) are not yet supported.</span></span>
- <span data-ttu-id="ea082-234">デバッグは、Excel for Windows の非同期関数に対してのみ有効です。</span><span class="sxs-lookup"><span data-stu-id="ea082-234">Debugging is only enabled for asynchronous functions on Excel for Windows.</span></span>
- <span data-ttu-id="ea082-235">Office 365 管理ポータルと AppSource による展開はまだ有効になっていません。</span><span class="sxs-lookup"><span data-stu-id="ea082-235">Deployment via the Office 365 Admin Portal and AppSource are not yet enabled.</span></span>
- <span data-ttu-id="ea082-236">Excel Online のカスタム関数は、一定期間使用しないとセッション中に機能しなくなることがあります。</span><span class="sxs-lookup"><span data-stu-id="ea082-236">Custom functions in Excel Online may stop working during a session after a period of inactivity.</span></span> <span data-ttu-id="ea082-237">ブラウザページを更新（F5）し、カスタム関数を再入力して機能を復元します。</span><span class="sxs-lookup"><span data-stu-id="ea082-237">Refresh the browser page (F5) and re-enter a custom function to restore the feature.</span></span>

## <a name="changelog"></a><span data-ttu-id="ea082-238">変更ログ</span><span class="sxs-lookup"><span data-stu-id="ea082-238">Changelog</span></span>

- <span data-ttu-id="ea082-239">**2017 年 11 月 7 日**: カスタム関数のプレビューとサンプルを公開\*</span><span class="sxs-lookup"><span data-stu-id="ea082-239">**Nov 7, 2017**: Shipped the custom functions preview and samples</span></span>
- <span data-ttu-id="ea082-240">**2017 年 11 月 20 日**: ビルド 8801 以降を使用する場合の互換性バグを修正</span><span class="sxs-lookup"><span data-stu-id="ea082-240">**Nov 20, 2017**: Fixed compatibility bug for those using builds 8801 and later</span></span>
- <span data-ttu-id="ea082-241">**2017 年 11 月 28 日**: 非同期関数のキャンセルのサポートを公開 (ストリーム機能の変更が必要)</span><span class="sxs-lookup"><span data-stu-id="ea082-241">**Nov 28, 2017**: Shipped support for cancellation on asynchronous functions (requires change for streaming functions)</span></span>
- <span data-ttu-id="ea082-242">**2018 年 5 月 7 日**：Mac、Excel Online、およびインプロセスで実行される同期関数のサポートを公開\*</span><span class="sxs-lookup"><span data-stu-id="ea082-242">**May 7, 2018**: Shipped support for Mac, Excel Online, and synchronous functions running in-process</span></span>

<span data-ttu-id="ea082-243">\* Office Insiders チャネル対象</span><span class="sxs-lookup"><span data-stu-id="ea082-243">\* to the Office Insiders Channel</span></span>
