---
title: Word アドインのチュートリアル
description: このチュートリアルでは、テキスト範囲、段落、画像、HTML、テーブル、コンテンツ コントロールを挿入 (および置換) する Word アドインを作成します。 テキストに書式を設定する方法と、コンテンツ コントロールにコンテンツを挿入 (および置換) する方法についても説明します。
ms.date: 06/20/2019
ms.prod: word
ms.topic: tutorial
localization_priority: Normal
ms.openlocfilehash: a9383128569a2cbe9b300ff9fee78d1dcb20e632
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/21/2019
ms.locfileid: "35126912"
---
# <a name="tutorial-create-a-word-task-pane-add-in"></a><span data-ttu-id="7aea3-104">チュートリアル: Word 作業ウィンドウ アドインを作成する</span><span class="sxs-lookup"><span data-stu-id="7aea3-104">Tutorial: Create a Word task pane add-in</span></span>

<span data-ttu-id="7aea3-105">このチュートリアルでは、以下の Word 作業ウィンドウ アドインを作成します。</span><span class="sxs-lookup"><span data-stu-id="7aea3-105">In this tutorial, you'll create a Word task pane add-in that:</span></span>

> [!div class="checklist"]
> * <span data-ttu-id="7aea3-106">テキスト範囲の挿入</span><span class="sxs-lookup"><span data-stu-id="7aea3-106">Inserts a range of text</span></span>
> * <span data-ttu-id="7aea3-107">テキストの書式設定</span><span class="sxs-lookup"><span data-stu-id="7aea3-107">Formats text</span></span>
> * <span data-ttu-id="7aea3-108">テキストの置換とさまざまな位置へのテキストの挿入</span><span class="sxs-lookup"><span data-stu-id="7aea3-108">Replaces text and inserts text in various locations</span></span>
> * <span data-ttu-id="7aea3-109">画像、HTML、テーブルの挿入</span><span class="sxs-lookup"><span data-stu-id="7aea3-109">Inserts images, HTML, and tables</span></span>
> * <span data-ttu-id="7aea3-110">コンテンツ コントロールの作成と更新</span><span class="sxs-lookup"><span data-stu-id="7aea3-110">Creates and updates content controls</span></span> 

## <a name="prerequisites"></a><span data-ttu-id="7aea3-111">前提条件</span><span class="sxs-lookup"><span data-stu-id="7aea3-111">Prerequisites</span></span>

<span data-ttu-id="7aea3-112">このチュートリアルを使用するには、以下のバージョンがインストールされている必要があります。</span><span class="sxs-lookup"><span data-stu-id="7aea3-112">To use this tutorial, you need to have the following installed.</span></span>

- <span data-ttu-id="7aea3-p102">Word 2016、バージョン 1711 (ビルド 8730.1000 クイック実行) 以降。 このバージョンを入手するには、Office Insider への参加が必要になることがあります。 詳細については、「[Office Insider](https://products.office.com/office-insider?tab=tab-1)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="7aea3-p102">Word 2016, version 1711 (Build 8730.1000 Click-to-Run) or later. You might need to be an Office Insider to get this version. For more information, see [Be an Office Insider](https://products.office.com/office-insider?tab=tab-1).</span></span>

- [<span data-ttu-id="7aea3-116">ノード</span><span class="sxs-lookup"><span data-stu-id="7aea3-116">Node</span></span>](https://nodejs.org/en/) 

- <span data-ttu-id="7aea3-117">[Git バッシュ](https://git-scm.com/downloads) (または別の Git クライアント)</span><span class="sxs-lookup"><span data-stu-id="7aea3-117">[Git Bash](https://git-scm.com/downloads) (or another Git client)</span></span>

## <a name="create-your-add-in-project"></a><span data-ttu-id="7aea3-118">アドイン プロジェクトの作成</span><span class="sxs-lookup"><span data-stu-id="7aea3-118">Create your add-in project</span></span>

<span data-ttu-id="7aea3-119">このチュートリアルの基礎として使用する Word アドイン プロジェクトを作成するには、次の手順を実行します。</span><span class="sxs-lookup"><span data-stu-id="7aea3-119">Complete the following steps to create the Word add-in project that you'll use as the basis for this tutorial.</span></span>

1. <span data-ttu-id="7aea3-120">「[Word アドインのチュートリアル](https://github.com/OfficeDev/Word-Add-in-Tutorial)」で、GitHub リポジトリを複製します。</span><span class="sxs-lookup"><span data-stu-id="7aea3-120">Clone the GitHub repository [Word add-in tutorial](https://github.com/OfficeDev/Word-Add-in-Tutorial).</span></span>

2. <span data-ttu-id="7aea3-121">Git bash ウィンドウまたは Node.JS 対応のシステム プロンプトを開いて、プロジェクトの **Start** フォルダーに移動します。</span><span class="sxs-lookup"><span data-stu-id="7aea3-121">Open a Git bash window, or Node.JS-enabled system prompt, and navigate to the **Start** folder of the project.</span></span>

3. <span data-ttu-id="7aea3-122">`npm install` コマンドを実行して、package.json ファイルに一覧表示されているツールとライブラリをインストールします。</span><span class="sxs-lookup"><span data-stu-id="7aea3-122">Run the command `npm install` to install the tools and libraries listed in the package.json file.</span></span> 

4. <span data-ttu-id="7aea3-123">開発用のコンピューターのオペレーティングシステムの証明書を信頼するように、[自己署名証明書をインストール](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md)する手順を実行します。</span><span class="sxs-lookup"><span data-stu-id="7aea3-123">Carry out the steps in [Installing the self-signed certificate](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md) to trust the certificate for your development computer's operating system.</span></span>

## <a name="insert-a-range-of-text"></a><span data-ttu-id="7aea3-124">テキスト範囲の挿入</span><span class="sxs-lookup"><span data-stu-id="7aea3-124">Insert a range of text</span></span>

<span data-ttu-id="7aea3-125">チュートリアルのこの手順では、ユーザーが現在使用している Word のバージョンをアドインがサポートしているかどうかをプログラムによってテストし、ドキュメントに段落を挿入します。</span><span class="sxs-lookup"><span data-stu-id="7aea3-125">In this step of the tutorial, you'll programmatically test that your add-in supports the user's current version of Word, and then insert a paragraph into the document.</span></span>

### <a name="code-the-add-in"></a><span data-ttu-id="7aea3-126">アドインのコードを作成する</span><span class="sxs-lookup"><span data-stu-id="7aea3-126">Code the add-in</span></span>

1. <span data-ttu-id="7aea3-127">コード エディターでプロジェクトを開きます。</span><span class="sxs-lookup"><span data-stu-id="7aea3-127">Open the project in your code editor.</span></span>

2. <span data-ttu-id="7aea3-128">index.html ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="7aea3-128">Open the file index.html.</span></span>

3. <span data-ttu-id="7aea3-129">`TODO1` を次のマークアップに置き換えます。</span><span class="sxs-lookup"><span data-stu-id="7aea3-129">Replace the `TODO1` with the following markup:</span></span>

    ```html
    <button class="ms-Button" id="insert-paragraph">Insert Paragraph</button>
    ```

4. <span data-ttu-id="7aea3-130">app.js ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="7aea3-130">Open the app.js file.</span></span>

5. <span data-ttu-id="7aea3-p103">`TODO1` を次のコードに置き換えます。 このコードでは、ユーザーの Word のバージョンが、このチュートリアルのすべての段階で使用するすべての API を含んでいる Word.js のバージョンをサポートしているかどうかを調べます。 運用アドインでは、未サポートの API を呼び出す UI を非表示または無効化する条件ブロックの本体を使用してください。 これにより、ユーザーは、自分が使用している Word のバージョンでサポートされているアドインの部分を使用できるようになります。</span><span class="sxs-lookup"><span data-stu-id="7aea3-p103">Replace the `TODO1` with the following code. This code determines whether the user's version of Word supports a version of Word.js that includes all the APIs that are used in all the stages of this tutorial. In a production add-in, use the body of the conditional block to hide or disable the UI that would call unsupported APIs. This will enable the user to still use the parts of the add-in that are supported by their version of Word.</span></span>

    ```js
    if (!Office.context.requirements.isSetSupported('WordApi', 1.3)) {
        console.log('Sorry. The tutorial add-in uses Word.js APIs that are not available in your version of Office.');
    }
    ```

6. <span data-ttu-id="7aea3-135">`TODO2` を次のコードに置き換えます。</span><span class="sxs-lookup"><span data-stu-id="7aea3-135">Replace the `TODO2` with the following code:</span></span>

    ```js
    $('#insert-paragraph').click(insertParagraph);
    ```

7. <span data-ttu-id="7aea3-136">`TODO3` を次のコードに置き換えます。</span><span class="sxs-lookup"><span data-stu-id="7aea3-136">Replace the `TODO3` with the following code.</span></span> <span data-ttu-id="7aea3-137">注:</span><span class="sxs-lookup"><span data-stu-id="7aea3-137">Note:</span></span>

   - <span data-ttu-id="7aea3-p105">Word.js のビジネス ロジックは、`Word.run` に渡される関数に追加されます。 このロジックは、すぐには実行されません。 その代わりに、保留中のコマンドのキューに追加されます。</span><span class="sxs-lookup"><span data-stu-id="7aea3-p105">Your Word.js business logic will be added to the function that is passed to `Word.run`. This logic does not execute immediately. Instead, it is added to a queue of pending commands.</span></span>

   - <span data-ttu-id="7aea3-141">`context.sync` メソッドは、キューに登録されたすべてのコマンドを、実行するために Word に送信します。</span><span class="sxs-lookup"><span data-stu-id="7aea3-141">The `context.sync` method sends all queued commands to Word for execution.</span></span>

   - <span data-ttu-id="7aea3-p106">`Word.run` の後に `catch` ブロックを続けます。 これは、どのような場合にも当てはまるベスト プラクティスです。</span><span class="sxs-lookup"><span data-stu-id="7aea3-p106">The `Word.run` is followed by a `catch` block. This is a best practice that you should always follow.</span></span> 

    ```js
    function insertParagraph() {
        Word.run(function (context) {

            // TODO4: Queue commands to insert a paragraph into the document.

            return context.sync();
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }
    ```

8. <span data-ttu-id="7aea3-144">`TODO4` を次のコードに置き換えます。</span><span class="sxs-lookup"><span data-stu-id="7aea3-144">Replace `TODO4` with the following code.</span></span> <span data-ttu-id="7aea3-145">次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="7aea3-145">Note:</span></span>

   - <span data-ttu-id="7aea3-146">`insertParagraph` メソッドの最初のパラメーターは、新しい段落のテキストです。</span><span class="sxs-lookup"><span data-stu-id="7aea3-146">The first parameter to the `insertParagraph` method is the text for the new paragraph.</span></span>

   - <span data-ttu-id="7aea3-p108">2 番目のパラメーターは、段落を挿入する本文内の場所です。 親オブジェクトが本文の場合、段落の挿入に使用できるその他のオプションには、End と Replace があります。</span><span class="sxs-lookup"><span data-stu-id="7aea3-p108">The second parameter is the location within the body where the paragraph will be inserted. Other options for insert paragraph, when the parent object is the body, are "End" and "Replace".</span></span>

    ```js
    var docBody = context.document.body;
    docBody.insertParagraph("Office has several versions, including Office 2016, Office 365 Click-to-Run, and Office on the web.",
                            "Start");
    ```

### <a name="test-the-add-in"></a><span data-ttu-id="7aea3-149">アドインをテストする</span><span class="sxs-lookup"><span data-stu-id="7aea3-149">Test the add-in</span></span>

1. <span data-ttu-id="7aea3-150">Git bash ウィンドウまたは Node.JS 対応のシステム プロンプトを開いて、プロジェクトの **Start** フォルダーに移動します。</span><span class="sxs-lookup"><span data-stu-id="7aea3-150">Open a Git bash window, or Node.JS-enabled system prompt, and navigate to the **Start** folder of the project.</span></span>

2. <span data-ttu-id="7aea3-151">`npm run build` コマンドを実行し、Office アドインを実行できるすべてのホストでサポートされている以前のバージョンの JavaScript に ES6 ソース コードをトランスパイルします。</span><span class="sxs-lookup"><span data-stu-id="7aea3-151">Run the command `npm run build` to transpile your ES6 source code to an earlier version of JavaScript that is supported by all the hosts where Office Add-ins can run.</span></span>

3. <span data-ttu-id="7aea3-152">`npm start` コマンドを実行して、ローカルホストで稼働する Web サーバーを起動します。</span><span class="sxs-lookup"><span data-stu-id="7aea3-152">Run the command `npm start` to start a web server running on localhost.</span></span>

4. <span data-ttu-id="7aea3-153">次のいずれかの方法を使用して、アドインをサイドロードします。</span><span class="sxs-lookup"><span data-stu-id="7aea3-153">Sideload the add-in by using one of the following methods:</span></span>

    - <span data-ttu-id="7aea3-154">Windows: [Windows で Office アドインをサイドロードする](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span><span class="sxs-lookup"><span data-stu-id="7aea3-154">Windows: [Sideload Office Add-ins on Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span></span>

    - <span data-ttu-id="7aea3-155">Web ブラウザー:[サイドロード Office アドイン (web)](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web)</span><span class="sxs-lookup"><span data-stu-id="7aea3-155">Web browser: [Sideload Office Add-ins in Office on the web](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web)</span></span>

    - <span data-ttu-id="7aea3-156">iPad および Mac: [iPad と Mac で Office アドインをサイドロードする](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span><span class="sxs-lookup"><span data-stu-id="7aea3-156">iPad and Mac: [Sideload Office Add-ins on iPad and Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span></span>

5. <span data-ttu-id="7aea3-157">Word の **[ホーム]** メニューで、**[作業ウィンドウの表示]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="7aea3-157">On the **Home** menu of Word, select **Show Taskpane**.</span></span>

6. <span data-ttu-id="7aea3-158">作業ウィンドウで、**[段落の挿入]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="7aea3-158">In the task pane, choose **Insert Paragraph**.</span></span>

7. <span data-ttu-id="7aea3-159">段落に変更を加えます。</span><span class="sxs-lookup"><span data-stu-id="7aea3-159">Make a change in the paragraph.</span></span>

8. <span data-ttu-id="7aea3-160">**[段落の挿入]** をもう一度選択します。</span><span class="sxs-lookup"><span data-stu-id="7aea3-160">Choose **Insert Paragraph** again.</span></span> <span data-ttu-id="7aea3-161">`insertParagraph` メソッドはドキュメントの本文の開始位置に挿入を行うため、新しい段落は前の段落より上に追加されます。</span><span class="sxs-lookup"><span data-stu-id="7aea3-161">Note that the new paragraph is above the previous one because the `insertParagraph` method is inserting at the start of the document's body.</span></span>

    ![Word のチュートリアル - 段落の挿入](../images/word-tutorial-insert-paragraph.png)

## <a name="format-text"></a><span data-ttu-id="7aea3-163">テキストの書式設定</span><span class="sxs-lookup"><span data-stu-id="7aea3-163">Format text</span></span>

<span data-ttu-id="7aea3-164">チュートリアルのこの手順では、組み込みのスタイルをテキストに適用したり、カスタム スタイルをテキストに適用したり、テキストのフォントを変更したりします。</span><span class="sxs-lookup"><span data-stu-id="7aea3-164">In this step of the tutorial, you'll apply a built-in style to text, apply a custom style to text, and change the font of text.</span></span>

### <a name="apply-a-built-in-style-to-text"></a><span data-ttu-id="7aea3-165">組み込みのスタイルをテキストに適用する</span><span class="sxs-lookup"><span data-stu-id="7aea3-165">Apply a built-in style to text</span></span>

1. <span data-ttu-id="7aea3-166">コード エディターでプロジェクトを開きます。</span><span class="sxs-lookup"><span data-stu-id="7aea3-166">Open the project in your code editor.</span></span> 

2. <span data-ttu-id="7aea3-167">index.html ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="7aea3-167">Open the file index.html.</span></span>

3. <span data-ttu-id="7aea3-168">`insert-paragraph` ボタンを格納している `div` の直下に、次のマークアップを追加します。</span><span class="sxs-lookup"><span data-stu-id="7aea3-168">Just below the `div` that contains the `insert-paragraph` button, add the following markup:</span></span>

    ```html
    <div class="padding">            
        <button class="ms-Button" id="apply-style">Apply Style</button>            
    </div>
    ```

4. <span data-ttu-id="7aea3-169">app.js ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="7aea3-169">Open the app.js file.</span></span>

5. <span data-ttu-id="7aea3-170">`insert-paragraph` ボタンにクリック ハンドラーを割り当てる行の直下に、次のコードを追加します。</span><span class="sxs-lookup"><span data-stu-id="7aea3-170">Just below the line that assigns a click handler to the `insert-paragraph` button, add the following code:</span></span>

    ```js
    $('#apply-style').click(applyStyle);
    ```

6. <span data-ttu-id="7aea3-171">`insertParagraph` 関数の直下に、次の関数を追加します。</span><span class="sxs-lookup"><span data-stu-id="7aea3-171">Just below the `insertParagraph` function, add the following function:</span></span>

    ```js
    function applyStyle() {
        Word.run(function (context) {
            
            // TODO1: Queue commands to style text.

            return context.sync();
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }
    ``` 

7. <span data-ttu-id="7aea3-p110">`TODO1` を次のコードに置き換えます。 このコードではスタイルを段落に適用していますが、スタイルはテキストの範囲にも適用できます。</span><span class="sxs-lookup"><span data-stu-id="7aea3-p110">Replace `TODO1` with the following code. Note that the code applies a style to a paragraph, but styles can also be applied to ranges of text.</span></span>

    ```js
    var firstParagraph = context.document.body.paragraphs.getFirst();
    firstParagraph.styleBuiltIn = Word.Style.intenseReference;
    ``` 

### <a name="apply-a-custom-style-to-text"></a><span data-ttu-id="7aea3-174">カスタム スタイルをテキストに適用する</span><span class="sxs-lookup"><span data-stu-id="7aea3-174">Apply a custom style to text</span></span>

1. <span data-ttu-id="7aea3-175">index.html ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="7aea3-175">Open the file index.html.</span></span>

2. <span data-ttu-id="7aea3-176">`apply-style` ボタンを格納している `div` の下に、次のマークアップを追加します。</span><span class="sxs-lookup"><span data-stu-id="7aea3-176">Below the `div` that contains the `apply-style` button, add the following markup:</span></span>

    ```html
    <div class="padding">            
        <button class="ms-Button" id="apply-custom-style">Apply Custom Style</button>            
    </div>
    ```

3. <span data-ttu-id="7aea3-177">app.js ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="7aea3-177">Open the app.js file.</span></span>

4. <span data-ttu-id="7aea3-178">`apply-style` ボタンにクリック ハンドラーを割り当てる行の下に、次のコードを追加します。</span><span class="sxs-lookup"><span data-stu-id="7aea3-178">Below the line that assigns a click handler to the `apply-style` button, add the following code:</span></span>

    ```js
    $('#apply-custom-style').click(applyCustomStyle);
    ```

5. <span data-ttu-id="7aea3-179">`applyStyle` 関数の下に、次の関数を追加します。</span><span class="sxs-lookup"><span data-stu-id="7aea3-179">Below the `applyStyle` function, add the following function:</span></span>

    ```js
    function applyCustomStyle() {
        Word.run(function (context) {
            
            // TODO1: Queue commands to apply the custom style.

            return context.sync();
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }
    ``` 

6. <span data-ttu-id="7aea3-p111">`TODO1` を次のコードに置き換えます。 このコードでは、まだ存在していないカスタム スタイルを適用しています。 「**アドインをテストする**」の手順で [MyCustomStyle](#test-the-add-in) という名前のスタイルを作成します。</span><span class="sxs-lookup"><span data-stu-id="7aea3-p111">Replace `TODO1` with the following code. Note that the code applies a custom style that does not exist yet. You'll create a style with the name **MyCustomStyle** in the [Test the add-in](#test-the-add-in) step.</span></span>

    ```js
    var lastParagraph = context.document.body.paragraphs.getLast();
    lastParagraph.style = "MyCustomStyle";
    ``` 

### <a name="change-the-font-of-text"></a><span data-ttu-id="7aea3-183">テキストのフォントを変更する</span><span class="sxs-lookup"><span data-stu-id="7aea3-183">Change the font of text</span></span>

1. <span data-ttu-id="7aea3-184">index.html ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="7aea3-184">Open the file index.html.</span></span>

2. <span data-ttu-id="7aea3-185">`apply-custom-style` ボタンを格納している `div` の下に、次のマークアップを追加します。</span><span class="sxs-lookup"><span data-stu-id="7aea3-185">Below the `div` that contains the `apply-custom-style` button, add the following markup:</span></span>

    ```html
    <div class="padding">            
        <button class="ms-Button" id="change-font">Change Font</button>            
    </div>
    ```

3. <span data-ttu-id="7aea3-186">app.js ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="7aea3-186">Open the app.js file.</span></span>

4. <span data-ttu-id="7aea3-187">`apply-custom-style` ボタンにクリック ハンドラーを割り当てる行の下に、次のコードを追加します。</span><span class="sxs-lookup"><span data-stu-id="7aea3-187">Below the line that assigns a click handler to the `apply-custom-style` button, add the following code:</span></span>

    ```js
    $('#change-font').click(changeFont);
    ```

5. <span data-ttu-id="7aea3-188">`applyCustomStyle` 関数の下に、次の関数を追加します。</span><span class="sxs-lookup"><span data-stu-id="7aea3-188">Below the `applyCustomStyle` function, add the following function:</span></span>

    ```js
    function changeFont() {
        Word.run(function (context) {
            
            // TODO1: Queue commands to apply a different font.

            return context.sync();
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }
    ``` 

6. <span data-ttu-id="7aea3-p112">`TODO1` を次のコードに置き換えます。 このコードでは、`ParagraphCollection.getFirst` メソッドにチェーンされた `Paragraph.getNext` メソッドを使用して 2 番目の段落への参照を取得することに注意してください。</span><span class="sxs-lookup"><span data-stu-id="7aea3-p112">Replace `TODO1` with the following code. Note that the code gets a reference to the second paragraph by using the `ParagraphCollection.getFirst` method chained to the `Paragraph.getNext` method.</span></span>

    ```js
    var secondParagraph = context.document.body.paragraphs.getFirst().getNext();
    secondParagraph.font.set({
            name: "Courier New",
            bold: true,
            size: 18
        });
    ``` 

### <a name="test-the-add-in"></a><span data-ttu-id="7aea3-191">アドインをテストする</span><span class="sxs-lookup"><span data-stu-id="7aea3-191">Test the add-in</span></span>

1. <span data-ttu-id="7aea3-192">Git bash ウィンドウまたは Node.JS 対応のシステム プロンプトが前の段階のチュートリアルから開いたままになっている場合は、Ctrl + C を 2 回入力して実行中の Web サーバーを停止します。</span><span class="sxs-lookup"><span data-stu-id="7aea3-192">In the Git bash window, or Node.JS-enabled system prompt, from the previous stage tutorial is still open, enter Ctrl+C twice to stop the running web server.</span></span> <span data-ttu-id="7aea3-193">それ以外の場合は、Git bash ウィンドウまたは Node.JS 対応のシステム プロンプトを開いて、プロジェクトの **Start** フォルダーに移動します。</span><span class="sxs-lookup"><span data-stu-id="7aea3-193">Otherwise, open a Git bash window, or Node.JS-enabled system prompt, and navigate to the **Start** folder of the project.</span></span>

     > [!NOTE]
     > <span data-ttu-id="7aea3-p114">ブラウザー同期サーバーは、app.js ファイルなどのファイルに変更を加えるたびに作業ウィンドウ内のアドインを再読み込みしますが、JavaScript を再トランスパイルしないため、ビルド コマンドを繰り返し実行して、app.js への変更を反映させる必要があります。 これを行うには、プロンプトが表示されてビルド コマンドを入力できるようにするため、サーバー プロセスを強制終了する必要があります。 ビルド後に、サーバーを再起動します。 次の数ステップで、このプロセスを実行します。</span><span class="sxs-lookup"><span data-stu-id="7aea3-p114">Although the browser-sync server reloads your add-in in the task pane every time you make a change to any file, including the app.js file, it does not retranspile the JavaScript, so you must repeat the build command in order for your changes to app.js to take effect. In order to do this, you need to kill the server process so that the prompt appears and you can enter the build command. After the build, you restart the server. The next few steps carry out this process.</span></span>

2. <span data-ttu-id="7aea3-198">`npm run build` コマンドを実行し、Office アドインを実行できるすべてのホストでサポートされている以前のバージョンの JavaScript に ES6 ソース コードをトランスパイルします。</span><span class="sxs-lookup"><span data-stu-id="7aea3-198">Run the command `npm run build` to transpile your ES6 source code to an earlier version of JavaScript that is supported by all the hosts where Office Add-ins can run.</span></span>

3. <span data-ttu-id="7aea3-199">`npm start` コマンドを実行して、ローカルホストで稼働する Web サーバーを起動します。</span><span class="sxs-lookup"><span data-stu-id="7aea3-199">Run the command `npm start` to start a web server running on localhost.</span></span>   

4. <span data-ttu-id="7aea3-200">作業ウィンドウを再読み込みするために、そのウィンドウを閉じて、**[ホーム]** メニューの **[作業ウィンドウの表示]** を選択してアドインを再度開きます。</span><span class="sxs-lookup"><span data-stu-id="7aea3-200">Reload the task pane by closing it, and then on the **Home** menu select **Show Taskpane** to reopen the add-in.</span></span>

5. <span data-ttu-id="7aea3-p115">ドキュメントに 3 つ以上の段落があることを確認してください。 **[段落の挿入]** を 3 回選択できます。 *ドキュメントの最後に空白の段落がないことを慎重にチェックしてください。空白の段落がある場合は、それを削除します。*</span><span class="sxs-lookup"><span data-stu-id="7aea3-p115">Be sure there are at least three paragraphs in the document. You can choose **Insert Paragraph** three times. *Check carefully that there's no blank paragraph at the end of the document. If there is, delete it.*</span></span>

6. <span data-ttu-id="7aea3-p116">Word で、MyCustomStyle という名前のカスタム スタイルを作成します。 このスタイルには、必要に応じて任意の書式を設定できます。</span><span class="sxs-lookup"><span data-stu-id="7aea3-p116">In Word, create a custom style named "MyCustomStyle". It can have any formatting that you want.</span></span>

7. <span data-ttu-id="7aea3-p117">**[スタイルの適用]** ボタンを選択します。 最初の段落は、組み込みのスタイルである **Intense Reference** でスタイル設定されます。</span><span class="sxs-lookup"><span data-stu-id="7aea3-p117">Choose the **Apply Style** button. The first paragraph will be styled with the built-in style **Intense Reference**.</span></span>

8. <span data-ttu-id="7aea3-p118">**[カスタム スタイルの適用]** ボタンを選択します。 最後の段落は、選択したカスタム スタイルでスタイル設定されます。 (何も起こらないように思える場合、最後の段落が空白である可能性があります。 その場合は、最後の段落にテキストを追加します)。</span><span class="sxs-lookup"><span data-stu-id="7aea3-p118">Choose the **Apply Custom Style** button. The last paragraph will be styled with your custom style. (If nothing seems to happen, the last paragraph might be blank. If so, add some text to it.)</span></span>

9. <span data-ttu-id="7aea3-p119">**[フォントの変更]** ボタンを選択します。 2 番目の段落のフォントを、18 ポイントで太字の Courier New に変更します。</span><span class="sxs-lookup"><span data-stu-id="7aea3-p119">Choose the **Change Font** button. The font of the second paragraph changes to 18 pt., bold, Courier New.</span></span>

    ![Word のチュートリアル - スタイルとフォントの適用](../images/word-tutorial-apply-styles-and-font.png)

## <a name="replace-text-and-insert-text"></a><span data-ttu-id="7aea3-215">テキストの置換と挿入</span><span class="sxs-lookup"><span data-stu-id="7aea3-215">Replace text and insert text</span></span>

<span data-ttu-id="7aea3-216">このチュートリアルの手順では、選択したテキスト範囲の内側や外側にテキストを追加したり、選択した範囲のテキストを置き換えたりします。</span><span class="sxs-lookup"><span data-stu-id="7aea3-216">In this step of the tutorial, you'll add text inside and outside of selected ranges of text, and replace the text of a selected range.</span></span>

### <a name="add-text-inside-a-range"></a><span data-ttu-id="7aea3-217">範囲内にテキストを追加する</span><span class="sxs-lookup"><span data-stu-id="7aea3-217">Add text inside a range</span></span>

1. <span data-ttu-id="7aea3-218">コード エディターでプロジェクトを開きます。</span><span class="sxs-lookup"><span data-stu-id="7aea3-218">Open the project in your code editor.</span></span>

2. <span data-ttu-id="7aea3-219">index.html ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="7aea3-219">Open the file index.html.</span></span>

3. <span data-ttu-id="7aea3-220">`change-font` ボタンを格納している `div` の下に、次のマークアップを追加します。</span><span class="sxs-lookup"><span data-stu-id="7aea3-220">Below the `div` that contains the `change-font` button, add the following markup:</span></span>

    ```html
    <div class="padding">
        <button class="ms-Button" id="insert-text-into-range">Insert Abbreviation</button>
    </div>
    ```

4. <span data-ttu-id="7aea3-221">app.js ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="7aea3-221">Open the app.js file.</span></span>

5. <span data-ttu-id="7aea3-222">`change-font` ボタンにクリック ハンドラーを割り当てる行の下に、次のコードを追加します。</span><span class="sxs-lookup"><span data-stu-id="7aea3-222">Below the line that assigns a click handler to the `change-font` button, add the following code:</span></span>

    ```js
    $('#insert-text-into-range').click(insertTextIntoRange);
    ```

6. <span data-ttu-id="7aea3-223">`changeFont` 関数の下に、次の関数を追加します。</span><span class="sxs-lookup"><span data-stu-id="7aea3-223">Below the `changeFont` function, add the following function:</span></span>

    ```js
    function insertTextIntoRange() {
        Word.run(function (context) {

            // TODO1: Queue commands to insert text into a selected range.

            // TODO2: Load the text of the range and sync so that the
            //        current range text can be read.

            // TODO3: Queue commands to repeat the text of the original
            //        range at the end of the document.

            return context.sync();
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }
    ``` 

7. <span data-ttu-id="7aea3-224">`TODO1` を次のコードに置き換えます。</span><span class="sxs-lookup"><span data-stu-id="7aea3-224">Replace `TODO1` with the following code.</span></span> <span data-ttu-id="7aea3-225">次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="7aea3-225">Note:</span></span>

   - <span data-ttu-id="7aea3-p121">このメソッドの目的は、テキストが Click-to-Run という範囲の末尾に (C2R) という省略形を挿入することです。 これは前提を単純化し、文字列は存在しており、ユーザーがその文字列を選択したものとしています。</span><span class="sxs-lookup"><span data-stu-id="7aea3-p121">The method is intended to insert the abbreviation ["(C2R)"] into the end of the Range whose text is "Click-to-Run". It makes a simplifying assumption that the string is present and the user has selected it.</span></span>

   - <span data-ttu-id="7aea3-228">`Range.insertText` メソッドの最初のパラメーターは、`Range` オブジェクトに挿入する文字列です。</span><span class="sxs-lookup"><span data-stu-id="7aea3-228">The first parameter of the `Range.insertText` method is the string to insert into the `Range` object.</span></span>

   - <span data-ttu-id="7aea3-p122">2 番目のパラメーターは、範囲内のどの位置にテキストを挿入するかを指定します。 End の他に、Start、Before、After、Replace が選択できます。</span><span class="sxs-lookup"><span data-stu-id="7aea3-p122">The second parameter specifies where in the range the additional text should be inserted. Besides "End", the other possible options are "Start", "Before", "After", and "Replace".</span></span> 

   - <span data-ttu-id="7aea3-p123">End と After の違いは、End が既存の範囲の内部の末尾に新しいテキストを挿入するのに対し、After の場合は文字列の入った新しい範囲を作成し、既存の範囲の後にその新しい範囲が挿入されることです。 同様に、Start はテキストを既存の範囲の内部の先頭に挿入しますが、Before の場合は新しい範囲を挿入します。 Replace は、既存の範囲のテキストを最初のパラメーターで指定した文字列に置き換えます。</span><span class="sxs-lookup"><span data-stu-id="7aea3-p123">The difference between "End" and "After" is that "End" inserts the new text inside the end of the existing range, but "After" creates a new range with the string and inserts the new range after the existing range. Similarly, "Start" inserts text inside the beginning of the existing range and "Before" inserts a new range. "Replace" replaces the text of the existing range with the string in the first parameter.</span></span>

   - <span data-ttu-id="7aea3-p124">チュートリアルの前の段階で示したとおり、ボディ オブジェクトの insert\* メソッドに Before オプションや After オプションはありません。 これは、文書の本文の外部にはコンテンツを挿入できないからです。</span><span class="sxs-lookup"><span data-stu-id="7aea3-p124">You saw in an earlier stage of the tutorial that the insert\* methods of the body object do not have the "Before" and "After" options. This is because you can't put content outside of the document's body.</span></span>

    ```js
    var doc = context.document;
    var originalRange = doc.getSelection();
    originalRange.insertText(" (C2R)", "End");
    ```

8. <span data-ttu-id="7aea3-p125">`TODO2` はスキップし、次のセクションに移ります。 `TODO3` を次のコードに置き換えます。 このコードは、このチュートリアルの最初の段階で作成したコードに似ていますが、文書の先頭ではなく末尾に新しい段落を挿入する点が異なります。 この新しい段落で、新しいテキストが元の範囲の一部になっていることが示されます。</span><span class="sxs-lookup"><span data-stu-id="7aea3-p125">We'll skip over `TODO2` until the next section. Replace `TODO3` with the following code. This code is similar to the code you created in the first stage of the tutorial, except that now you are inserting a new paragraph at the end of the document instead of at the start. This new paragraph will demonstrate that the new text is now part of the original range.</span></span>

    ```js
    doc.body.insertParagraph("Original range: " + originalRange.text, "End");
    ```

### <a name="add-code-to-fetch-document-properties-into-the-task-panes-script-objects"></a><span data-ttu-id="7aea3-240">ドキュメントのプロパティを作業ウィンドウのスクリプト オブジェクトにフェッチするコードを追加する</span><span class="sxs-lookup"><span data-stu-id="7aea3-240">Add code to fetch document properties into the task pane's script objects</span></span>

<span data-ttu-id="7aea3-p126">このチュートリアルのシリーズで前述したすべての関数では、Office ドキュメントへの*書き込み*コマンドをキューに登録していました。 各関数は、キューに登録されたコマンドを実行対象のドキュメントに送信する `context.sync()` メソッドを呼び出すことで終了しています。 ただし、最後の手順で追加したコードでは、`originalRange.text` プロパティを呼び出しています。このことが、これまでに作成した関数とは大きく異なります。`originalRange` オブジェクトは、この作業ウィンドウのスクリプトに存在する単なるプロキシ オブジェクトなので、 ドキュメントの指定された範囲にある実際のテキストを認識できません。そのため、その `text` プロパティでは実際の値が保持できません。 まず、ドキュメントからその範囲のテキスト値をフェッチする必要があり、その値を使用して `originalRange.text` の値を設定します。 そのようにした場合にのみ、例外がスローされることなく `originalRange.text` を呼び出せるようになります。 このフェッチ処理には、3 つの手順があります。</span><span class="sxs-lookup"><span data-stu-id="7aea3-p126">In all the previous functions in this series of tutorials, you queued commands to *write* to the Office document. Each function ended with a call to the `context.sync()` method which sends the queued commands to the document to be executed. But the code you added in the last step calls the `originalRange.text` property, and this is a significant difference from the earlier functions you wrote, because the `originalRange` object is only a proxy object that exists in your task pane's script. It doesn't know what the actual text of the range in the document is, so its `text` property can't have a real value. It is necessary to first fetch the text value of the range from the document and use it to set the value of `originalRange.text`. Only then can `originalRange.text` be called without causing an exception to be thrown. This fetching process has three steps:</span></span>

   1. <span data-ttu-id="7aea3-248">コードで読み取る必要があるプロパティをロードする (つまりフェッチする) コマンドをキューに登録します。</span><span class="sxs-lookup"><span data-stu-id="7aea3-248">Queue a command to load (that is; fetch) the properties that your code needs to read.</span></span>

   2. <span data-ttu-id="7aea3-249">コンテキスト オブジェクトの `sync` メソッドを呼び出します。このメソッドは、キューに登録されたコマンドを実行対象のドキュメントに送信して、要求された情報を返します。</span><span class="sxs-lookup"><span data-stu-id="7aea3-249">Call the context object's `sync` method to send the queued command to the document for execution and return the requested information.</span></span>

   3. <span data-ttu-id="7aea3-250">`sync` メソッドは非同期であるため、フェッチされたプロパティをコードで呼び出す前に、そのメソッドが完了していることを確認します。</span><span class="sxs-lookup"><span data-stu-id="7aea3-250">Because the `sync` method is asynchronous, ensure that it has completed before your code calls the properties that were fetched.</span></span>

<span data-ttu-id="7aea3-251">こうした手順は、コードで Office ドキュメントから情報を*読み取る*必要がある場合には必ず完了する必要があります。</span><span class="sxs-lookup"><span data-stu-id="7aea3-251">These steps must be completed whenever your code needs to *read* information from the Office document.</span></span>

1. <span data-ttu-id="7aea3-252">`TODO2` を次のコードに置き換えます。</span><span class="sxs-lookup"><span data-stu-id="7aea3-252">Replace `TODO2` with the following code.</span></span>
  
    ```js
    originalRange.load("text");
    return context.sync()
        .then(function() {

                // TODO4: Move the doc.body.insertParagraph line here.

            }
        )
            // TODO5: Move the final call of context.sync here and ensure
            //        that it does not run until the insertParagraph has
            //        been queued.
    ```

2. <span data-ttu-id="7aea3-p127">分岐していない同一のコード パスに 2 つの `return` ステートメントを含めることはできないため、`Word.run` の最後にある最終行の `return context.sync();` を削除します。新しい最後の `context.sync` は、このチュートリアルの後の方で追加します。</span><span class="sxs-lookup"><span data-stu-id="7aea3-p127">You can't have two `return` statements in the same unbranching code path, so delete the final line `return context.sync();` at the end of the `Word.run`. You'll add a new final `context.sync` later in this tutorial.</span></span>

3. <span data-ttu-id="7aea3-255">`doc.body.insertParagraph` 行を切り取り、`TODO4` の代わりに貼り付けます。</span><span class="sxs-lookup"><span data-stu-id="7aea3-255">Cut the `doc.body.insertParagraph` line and paste in place of `TODO4`.</span></span>

4. <span data-ttu-id="7aea3-p128">`TODO5` を次のコードに置き換えます。次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="7aea3-p128">Replace `TODO5` with the following code. Note:</span></span>

   - <span data-ttu-id="7aea3-258">`sync` メソッドを `then` 関数に渡すことで、`insertParagraph` ロジックがキューに登録されるまで、そのメソッドが実行されないようにします。</span><span class="sxs-lookup"><span data-stu-id="7aea3-258">Passing the `sync` method to a `then` function ensures that it does not run until the `insertParagraph` logic has been queued.</span></span>

   - <span data-ttu-id="7aea3-259">`then` メソッドは渡されたどんな関数でも呼び出します。`sync` が 2 回呼び出されないように、context.sync の末尾の "()" は省略します。</span><span class="sxs-lookup"><span data-stu-id="7aea3-259">The `then` method invokes whatever function is passed to it, and you don't want `sync` to be invoked twice, so leave off the "()" from the end of context.sync.</span></span>

    ```js
    .then(context.sync);
    ```

<span data-ttu-id="7aea3-260">作業が完了すると、関数の全体は次のようになります。</span><span class="sxs-lookup"><span data-stu-id="7aea3-260">When you're done, the entire function should look like the following:</span></span>

```js
function insertTextIntoRange() {
    Word.run(function (context) {

        var doc = context.document;
        var originalRange = doc.getSelection();
        originalRange.insertText(" (C2R)", "End");

        originalRange.load("text");
        return context.sync()
            .then(function() {
                        doc.body.insertParagraph("Current text of original range: " + originalRange.text,
                                                "End");
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
}
```

### <a name="add-text-between-ranges"></a><span data-ttu-id="7aea3-261">範囲間にテキストを追加する</span><span class="sxs-lookup"><span data-stu-id="7aea3-261">Add text between ranges</span></span>

1. <span data-ttu-id="7aea3-262">index.html ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="7aea3-262">Open the file index.html.</span></span>

2. <span data-ttu-id="7aea3-263">`insert-text-into-range` ボタンを格納している `div` の下に、次のマークアップを追加します。</span><span class="sxs-lookup"><span data-stu-id="7aea3-263">Below the `div` that contains the `insert-text-into-range` button, add the following markup:</span></span>

    ```html
    <div class="padding">
        <button class="ms-Button" id="insert-text-outside-range">Add Version Info</button>
    </div>
    ```

3. <span data-ttu-id="7aea3-264">app.js ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="7aea3-264">Open the app.js file.</span></span>

4. <span data-ttu-id="7aea3-265">`insert-text-into-range` ボタンにクリック ハンドラーを割り当てる行の下に、次のコードを追加します。</span><span class="sxs-lookup"><span data-stu-id="7aea3-265">Below the line that assigns a click handler to the `insert-text-into-range` button, add the following code:</span></span>

    ```js
    $('#insert-text-outside-range').click(insertTextBeforeRange);
    ```

5. <span data-ttu-id="7aea3-266">`insertTextIntoRange` 関数の下に、次の関数を追加します。</span><span class="sxs-lookup"><span data-stu-id="7aea3-266">Below the `insertTextIntoRange` function, add the following function:</span></span>

    ```js
    function insertTextBeforeRange() {
        Word.run(function (context) {

            // TODO1: Queue commands to insert a new range before the
            //        selected range.

            // TODO2: Load the text of the original range and sync so that the
            //        range text can be read and inserted.

        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }
    ```

6. <span data-ttu-id="7aea3-267">`TODO1` を次のコードに置き換えます。</span><span class="sxs-lookup"><span data-stu-id="7aea3-267">Replace `TODO1` with the following code.</span></span> <span data-ttu-id="7aea3-268">次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="7aea3-268">Note:</span></span>

   - <span data-ttu-id="7aea3-p130">このメソッドの目的は、Office 365 というテキストから成る範囲の前に Office 2019 というテキストの範囲を追加することです。 これは前提を単純化し、文字列は存在しており、ユーザーがその文字列を選択したものとしています。</span><span class="sxs-lookup"><span data-stu-id="7aea3-p130">The method is intended to add a range whose text is "Office 2019, " before the range with text "Office 365". It makes a simplifying assumption that the string is present and the user has selected it.</span></span>

   - <span data-ttu-id="7aea3-271">`Range.insertText` メソッドの最初のパラメーターは、追加する文字列です。</span><span class="sxs-lookup"><span data-stu-id="7aea3-271">The first parameter of the `Range.insertText` method is the string to add.</span></span>

   - <span data-ttu-id="7aea3-p131">2 番目のパラメーターは、範囲内のどの位置にテキストを挿入するかを指定します。 位置オプションの詳細については、`insertTextIntoRange` 関数に関する上記の説明を参照してください。</span><span class="sxs-lookup"><span data-stu-id="7aea3-p131">The second parameter specifies where in the range the additional text should be inserted. For more details about the location options, see the previous discussion of the `insertTextIntoRange` function.</span></span>

    ```js
    var doc = context.document;
    var originalRange = doc.getSelection();
    originalRange.insertText("Office 2019, ", "Before");
    ```

7. <span data-ttu-id="7aea3-274">`TODO2` を次のコードに置き換えます。</span><span class="sxs-lookup"><span data-stu-id="7aea3-274">Replace `TODO2` with the following code.</span></span>

     ```js
    originalRange.load("text");
    return context.sync()
        .then(function() {

                // TODO3: Queue commands to insert the original range as a
                //        paragraph at the end of the document.

                }
            )

            // TODO4: Make a final call of context.sync here and ensure
            //        that it does not run until the insertParagraph has
            //        been queued.
    ```

8. <span data-ttu-id="7aea3-p132">`TODO3` を次のコードに置き換えます。 この新しい段落で、新しいテキストが元の選択範囲の一部になって***いない***ことが示されます。 元の範囲には、依然として選択時のテキストのみが含まれています。</span><span class="sxs-lookup"><span data-stu-id="7aea3-p132">Replace `TODO3` with the following code. This new paragraph will demonstrate the fact that the new text is ***not*** part of the original selected range. The original range still has only the text it had when it was selected.</span></span>

    ```js
    doc.body.insertParagraph("Current text of original range: " + originalRange.text,
                             "End");
    ```

9. <span data-ttu-id="7aea3-278">`TODO4` を次のコードに置き換えます。</span><span class="sxs-lookup"><span data-stu-id="7aea3-278">Replace `TODO4` with the following code:</span></span>

    ```js
    .then(context.sync);
    ```

### <a name="replace-the-text-of-a-range"></a><span data-ttu-id="7aea3-279">範囲のテキストを置き換える</span><span class="sxs-lookup"><span data-stu-id="7aea3-279">Replace the text of a range</span></span>

1. <span data-ttu-id="7aea3-280">index.html ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="7aea3-280">Open the file index.html.</span></span>

2. <span data-ttu-id="7aea3-281">`insert-text-outside-range` ボタンを格納している `div` の下に、次のマークアップを追加します。</span><span class="sxs-lookup"><span data-stu-id="7aea3-281">Below the `div` that contains the `insert-text-outside-range` button, add the following markup:</span></span>

    ```html
    <div class="padding">
        <button class="ms-Button" id="replace-text">Change Quantity Term</button>
    </div>
    ```

3. <span data-ttu-id="7aea3-282">app.js ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="7aea3-282">Open the app.js file.</span></span>

4. <span data-ttu-id="7aea3-283">`insert-text-outside-range` ボタンにクリック ハンドラーを割り当てる行の下に、次のコードを追加します。</span><span class="sxs-lookup"><span data-stu-id="7aea3-283">Below the line that assigns a click handler to the `insert-text-outside-range` button, add the following code:</span></span>

    ```js
    $('#replace-text').click(replaceText);
    ```

5. <span data-ttu-id="7aea3-284">`insertTextBeforeRange` 関数の下に、次の関数を追加します。</span><span class="sxs-lookup"><span data-stu-id="7aea3-284">Below the `insertTextBeforeRange` function, add the following function:</span></span>

    ```js
    function replaceText() {
        Word.run(function (context) {

            // TODO1: Queue commands to replace the text.

            return context.sync();
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }
    ```

6. <span data-ttu-id="7aea3-p133">`TODO1` を次のコードに置き換えます。 このメソッドの目的は、several という文字列を many という文字列で置き換えることです。 これは前提を単純化し、文字列は存在しており、ユーザーがその文字列を選択したものとしています。</span><span class="sxs-lookup"><span data-stu-id="7aea3-p133">Replace `TODO1` with the following code. Note that the method is intended to replace the string "several" with the string "many". It makes a simplifying assumption that the string is present and the user has selected it.</span></span>

    ```js
    var doc = context.document;
    var originalRange = doc.getSelection();
    originalRange.insertText("many", "Replace");
    ```

### <a name="test-the-add-in"></a><span data-ttu-id="7aea3-288">アドインをテストする</span><span class="sxs-lookup"><span data-stu-id="7aea3-288">Test the add-in</span></span>

1. <span data-ttu-id="7aea3-p134">Git bash ウィンドウまたは Node.JS 対応のシステム プロンプトが前の段階のチュートリアルから開いたままになっている場合は、Ctrl-C を 2 回入力して実行中の Web サーバーを停止します。 それ以外の場合は、Git bash ウィンドウまたは Node.JS 対応のシステム プロンプトを開いて、プロジェクトの **Start** フォルダーに移動します。</span><span class="sxs-lookup"><span data-stu-id="7aea3-p134">If the Git bash window, or Node.JS-enabled system prompt, from the previous stage tutorial is still open, enter Ctrl-C twice to stop the running web server. Otherwise, open a Git bash window, or Node.JS-enabled system prompt, and navigate to the **Start** folder of the project.</span></span>

     > [!NOTE]
     > <span data-ttu-id="7aea3-p135">ブラウザー同期サーバーは、app.js ファイルなどのファイルに変更を加えるたびに作業ウィンドウ内のアドインを再読み込みしますが、JavaScript を再トランスパイルしないため、ビルド コマンドを繰り返し実行して、app.js への変更を反映させる必要があります。 これを行うには、プロンプトが表示されてビルド コマンドを入力できるようにするため、サーバー プロセスを強制終了する必要があります。 ビルド後に、サーバーを再起動します。 次の数ステップで、このプロセスを実行します。</span><span class="sxs-lookup"><span data-stu-id="7aea3-p135">Although the browser-sync server reloads your add-in in the task pane every time you make a change to any file, including the app.js file, it does not retranspile the JavaScript, so you must repeat the build command in order for your changes to app.js to take effect. In order to do this, you need to kill the server process so that the prompt appears and you can enter the build command. After the build, restart the server. The next few steps carry out this process.</span></span>

2. <span data-ttu-id="7aea3-295">`npm run build` コマンドを実行し、Office アドインを実行できるすべてのホストでサポートされている以前のバージョンの JavaScript に ES6 ソース コードをトランスパイルします。</span><span class="sxs-lookup"><span data-stu-id="7aea3-295">Run the command `npm run build` to transpile your ES6 source code to an earlier version of JavaScript that is supported by all the hosts where Office Add-ins can run.</span></span>

3. <span data-ttu-id="7aea3-296">`npm start` コマンドを実行して、ローカルホストで稼働する Web サーバーを起動します。</span><span class="sxs-lookup"><span data-stu-id="7aea3-296">Run the command `npm start` to start a web server running on localhost.</span></span>

4. <span data-ttu-id="7aea3-297">作業ウィンドウを再読み込みするために、そのウィンドウを閉じて、**[ホーム]** メニューの **[作業ウィンドウの表示]** を選択してアドインを再度開きます。</span><span class="sxs-lookup"><span data-stu-id="7aea3-297">Reload the task pane by closing it, and then on the **Home** menu, select **Show Taskpane** to reopen the add-in.</span></span>

5. <span data-ttu-id="7aea3-298">作業ウィンドウで **[段落の挿入]** を選択し、文書の先頭に段落があることを確認します。</span><span class="sxs-lookup"><span data-stu-id="7aea3-298">In the task pane, choose **Insert Paragraph** to ensure that there is a paragraph at the start of the document.</span></span>

6. <span data-ttu-id="7aea3-p136">一部のテキストを選択します。 Click-to-Run という語句を選択します。 *選択範囲の前後にあるスペースは含めないように注意してください。*</span><span class="sxs-lookup"><span data-stu-id="7aea3-p136">Select some text. Selecting the phrase "Click-to-Run" will make the most sense. *Be careful not to include the preceding or following space in the selection.*</span></span>

7. <span data-ttu-id="7aea3-p137">**[ラベル (短縮形) の挿入]** ボタンを選択します。 (C2R) が追加されることに注意してください。 また、この新しい文字列は既存の範囲に追加されるため、文書の下部に新しい段落が追加され、拡張されたテキスト全体が含まれていることに注意してください。</span><span class="sxs-lookup"><span data-stu-id="7aea3-p137">Choose the **Insert Abbreviation** button. Note that " (C2R)" is added. Note also that at the bottom of the document a new paragraph is added with the entire expanded text because the new string was added to the existing range.</span></span>

8. <span data-ttu-id="7aea3-p138">一部のテキストを選択します。 Office 365 という語句を選択します。 *選択範囲の前後にあるスペースは含めないように注意してください。*</span><span class="sxs-lookup"><span data-stu-id="7aea3-p138">Select some text. Selecting the phrase "Office 365" will make the most sense. *Be careful not to include the preceding or following space in the selection.*</span></span>

9. <span data-ttu-id="7aea3-p139">**[バージョン情報の追加]** ボタンを選択します。 Office 2019 が、Office 2016 と Office 365 の間に挿入されることに注意してください。 また、この新しい文字列は元の範囲に追加されるのではなく新しい範囲になるため、文書の下部に新しい段落が追加されるものの、その段落には最初に選択したテキストのみが含まれることに注意してください。</span><span class="sxs-lookup"><span data-stu-id="7aea3-p139">Choose the **Add Version Info** button. Note that "Office 2019, " is inserted between "Office 2016" and "Office 365". Note also that at the bottom of the document a new paragraph is added but it contains only the originally selected text because the new string became a new range rather than being added to the original range.</span></span>

10. <span data-ttu-id="7aea3-p140">一部のテキストを選択します。 several という語句を選択します。 *選択範囲の前後にあるスペースは含めないように注意してください。*</span><span class="sxs-lookup"><span data-stu-id="7aea3-p140">Select some text. Selecting the word "several" will make the most sense. *Be careful not to include the preceding or following space in the selection.*</span></span>

11. <span data-ttu-id="7aea3-p141">**[数量の用語の変更]** ボタンを選択します。 選択したテキストが many に置き換えられることに注意してください。</span><span class="sxs-lookup"><span data-stu-id="7aea3-p141">Choose the **Change Quantity Term** button. Note that "many" replaces the selected text.</span></span>

    ![Word のチュートリアル - テキストの追加と置換](../images/word-tutorial-text-replace.png)

## <a name="insert-images-html-and-tables"></a><span data-ttu-id="7aea3-317">画像、HTML、テーブルの挿入</span><span class="sxs-lookup"><span data-stu-id="7aea3-317">Insert images, HTML, and tables</span></span>

<span data-ttu-id="7aea3-318">チュートリアルのこの手順では、ドキュメントに画像、HTML、テーブルを挿入する方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="7aea3-318">In this step of the tutorial, you'll learn how to insert images, HTML, and tables into the document.</span></span>

### <a name="insert-an-image"></a><span data-ttu-id="7aea3-319">画像の挿入</span><span class="sxs-lookup"><span data-stu-id="7aea3-319">Insert an image</span></span>

1. <span data-ttu-id="7aea3-320">コード エディターでプロジェクトを開きます。</span><span class="sxs-lookup"><span data-stu-id="7aea3-320">Open the project in your code editor.</span></span>

2. <span data-ttu-id="7aea3-321">index.html ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="7aea3-321">Open the file index.html.</span></span>

3. <span data-ttu-id="7aea3-322">`replace-text` ボタンを格納している `div` の下に、次のマークアップを追加します。</span><span class="sxs-lookup"><span data-stu-id="7aea3-322">Below the `div` that contains the `replace-text` button, add the following markup:</span></span>

    ```html
    <div class="padding">
        <button class="ms-Button" id="insert-image">Insert Image</button>
    </div>
    ```

4. <span data-ttu-id="7aea3-323">app.js ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="7aea3-323">Open the app.js file.</span></span>

5. <span data-ttu-id="7aea3-p142">ファイルの先頭近くにある、use-strict 行のすぐ下に次の行を追加します。 この行は、別のファイルから変数をインポートします。 この変数は、画像をエンコードする Base 64 文字列です。 エンコードされた文字列を表示するには、プロジェクトのルートにある base64Image.js ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="7aea3-p142">Near the top of the file, just below the use-strict line, add the following line. This line imports a variable from another file. The variable is a base 64 string that encodes an image. To see the encoded string, open the base64Image.js file in the root of the project.</span></span>

    ```js
    import { base64Image } from "./base64Image";
    ```

6. <span data-ttu-id="7aea3-328">`replace-text` ボタンにクリック ハンドラーを割り当てる行の下に、次のコードを追加します。</span><span class="sxs-lookup"><span data-stu-id="7aea3-328">Below the line that assigns a click handler to the `replace-text` button, add the following code:</span></span>

    ```js
    $('#insert-image').click(insertImage);
    ```

7. <span data-ttu-id="7aea3-329">`replaceText` 関数の下に、次の関数を追加します。</span><span class="sxs-lookup"><span data-stu-id="7aea3-329">Below the `replaceText` function, add the following function:</span></span>

    ```js
    function insertImage() {
        Word.run(function (context) {

            // TODO1: Queue commands to insert an image.

            return context.sync();
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }
    ```

8. <span data-ttu-id="7aea3-p143">`TODO1` を次のコードに置き換えます。 この行により、Base 64 でエンコードされた画像がドキュメントの末尾に挿入されることに注意してください。 (`Paragraph` オブジェクトにも `insertInlinePictureFromBase64` メソッドやその他の `insert*` メソッドがあります。 例については、次の insertHTML セクションを参照してください)。</span><span class="sxs-lookup"><span data-stu-id="7aea3-p143">Replace `TODO1` with the following code. Note that this line inserts the base 64 encoded image at the end of the document. (The `Paragraph` object also has an `insertInlinePictureFromBase64` method and other `insert*` methods. See the following insertHTML section for an example.)</span></span>

    ```js
    context.document.body.insertInlinePictureFromBase64(base64Image, "End");
    ```

### <a name="insert-html"></a><span data-ttu-id="7aea3-334">HTML の挿入</span><span class="sxs-lookup"><span data-stu-id="7aea3-334">Insert HTML</span></span>

1. <span data-ttu-id="7aea3-335">index.html ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="7aea3-335">Open the file index.html.</span></span>

2. <span data-ttu-id="7aea3-336">`insert-image` ボタンを格納している `div` の下に、次のマークアップを追加します。</span><span class="sxs-lookup"><span data-stu-id="7aea3-336">Below the `div` that contains the `insert-image` button, add the following markup:</span></span>

    ```html
    <div class="padding">
        <button class="ms-Button" id="insert-html">Insert HTML</button>
    </div>
    ```

3. <span data-ttu-id="7aea3-337">app.js ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="7aea3-337">Open the app.js file.</span></span>

4. <span data-ttu-id="7aea3-338">`insert-image` ボタンにクリック ハンドラーを割り当てる行の下に、次のコードを追加します。</span><span class="sxs-lookup"><span data-stu-id="7aea3-338">Below the line that assigns a click handler to the `insert-image` button, add the following code:</span></span>

    ```js
    $('#insert-html').click(insertHTML);
    ```

5. <span data-ttu-id="7aea3-339">`insertImage` 関数の下に、次の関数を追加します。</span><span class="sxs-lookup"><span data-stu-id="7aea3-339">Below the `insertImage` function, add the following function:</span></span>

    ```js
    function insertHTML() {
        Word.run(function (context) {

            // TODO1: Queue commands to insert a string of HTML.

            return context.sync();
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }
    ```

6. <span data-ttu-id="7aea3-340">`TODO1` を次のコードに置き換えます。</span><span class="sxs-lookup"><span data-stu-id="7aea3-340">Replace `TODO1` with the following code.</span></span> <span data-ttu-id="7aea3-341">次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="7aea3-341">Note:</span></span>

   - <span data-ttu-id="7aea3-342">最初の行は、ドキュメントの末尾に空白の段落を追加します。</span><span class="sxs-lookup"><span data-stu-id="7aea3-342">The first line adds a blank paragraph to the end of the document.</span></span> 

   - <span data-ttu-id="7aea3-p145">2 行目は、その段落の末尾に HTML の文字列を挿入します。具体的には、Verdana フォントで書式設定された段落と、Word 文書の既定のスタイルが設定された段落の 2 つの段落が挿入されます。 (`insertImage` メソッドで説明したように、`context.document.body` オブジェクトにも `insert*` メソッドがあります)。</span><span class="sxs-lookup"><span data-stu-id="7aea3-p145">The second line inserts a string of HTML at the end of the paragraph; specifically two paragraphs, one formatted with Verdana font, the other with the default styling of the Word document. (As you saw in the `insertImage` method earlier, the `context.document.body` object also has the `insert*` methods.)</span></span>

    ```js
    var blankParagraph = context.document.body.paragraphs.getLast().insertParagraph("", "After");
    blankParagraph.insertHtml('<p style="font-family: verdana;">Inserted HTML.</p><p>Another paragraph</p>', "End");
    ```

### <a name="insert-a-table"></a><span data-ttu-id="7aea3-345">テーブルの挿入</span><span class="sxs-lookup"><span data-stu-id="7aea3-345">Insert a table</span></span>

1. <span data-ttu-id="7aea3-346">index.html ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="7aea3-346">Open the file index.html.</span></span>

2. <span data-ttu-id="7aea3-347">`insert-html` ボタンを格納している `div` の下に、次のマークアップを追加します。</span><span class="sxs-lookup"><span data-stu-id="7aea3-347">Below the `div` that contains the `insert-html` button, add the following markup:</span></span>

    ```html
    <div class="padding">
        <button class="ms-Button" id="insert-table">Insert Table</button>
    </div>
    ```

3. <span data-ttu-id="7aea3-348">app.js ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="7aea3-348">Open the app.js file.</span></span>

4. <span data-ttu-id="7aea3-349">`insert-html` ボタンにクリック ハンドラーを割り当てる行の下に、次のコードを追加します。</span><span class="sxs-lookup"><span data-stu-id="7aea3-349">Below the line that assigns a click handler to the `insert-html` button, add the following code:</span></span>

    ```js
    $('#insert-table').click(insertTable);
    ```

5. <span data-ttu-id="7aea3-350">`insertHTML` 関数の下に、次の関数を追加します。</span><span class="sxs-lookup"><span data-stu-id="7aea3-350">Below the `insertHTML` function, add the following function:</span></span>

    ```js
    function insertTable() {
        Word.run(function (context) {

            // TODO1: Queue commands to get a reference to the paragraph
            //        that will proceed the table.

            // TODO2: Queue commands to create a table and populate it with data.

            return context.sync();
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }
    ```

6. <span data-ttu-id="7aea3-p146">`TODO1` を次のコードに置き換えます。 この行は `ParagraphCollection.getFirst` メソッドを使用して最初の段落への参照を取得し、次に `Paragraph.getNext` メソッドを使用して 2 番目の段落への参照を取得することに注意してください。</span><span class="sxs-lookup"><span data-stu-id="7aea3-p146">Replace `TODO1` with the following code. Note that this line uses the `ParagraphCollection.getFirst` method to get a reference ot the first paragraph and then uses the `Paragraph.getNext` method to get a reference to the second paragraph.</span></span>

    ```js
    var secondParagraph = context.document.body.paragraphs.getFirst().getNext();
    ```

7. <span data-ttu-id="7aea3-353">`TODO2` を次のコードに置き換えます。</span><span class="sxs-lookup"><span data-stu-id="7aea3-353">Replace `TODO2` with the following code.</span></span> <span data-ttu-id="7aea3-354">次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="7aea3-354">Note:</span></span>

   - <span data-ttu-id="7aea3-355">`insertTable` メソッドの最初の 2 つのパラメーターは、行と列の数を指定します。</span><span class="sxs-lookup"><span data-stu-id="7aea3-355">The first two parameters of the `insertTable` method specify the number of rows and columns.</span></span>

   - <span data-ttu-id="7aea3-356">3 番目のパラメーターは、テーブルを挿入する場所を指定します (この例では段落の後)。</span><span class="sxs-lookup"><span data-stu-id="7aea3-356">The third parameter specifies where to insert the table, in this case after the paragraph.</span></span>

   - <span data-ttu-id="7aea3-357">4 番目のパラメーターは、テーブルのセルの値を設定する 2 次元配列です。</span><span class="sxs-lookup"><span data-stu-id="7aea3-357">The fourth parameter is a two-dimensional array that sets the values of the table cells.</span></span>

   - <span data-ttu-id="7aea3-358">このテーブルには既定のスタイルがそのまま設定されますが、`insertTable` メソッドがさまざまなメンバーを持つ `Table` オブジェクトを返し、その一部がテーブルのスタイル設定に使用されます。</span><span class="sxs-lookup"><span data-stu-id="7aea3-358">The table will have plain default styling, but the `insertTable` method returns a `Table` object with many members, some of which are used to style the table.</span></span>

    ```js
    var tableData = [
            ["Name", "ID", "Birth City"],
            ["Bob", "434", "Chicago"],
            ["Sue", "719", "Havana"],
        ];
    secondParagraph.insertTable(3, 3, "After", tableData);
    ```

### <a name="test-the-add-in"></a><span data-ttu-id="7aea3-359">アドインをテストする</span><span class="sxs-lookup"><span data-stu-id="7aea3-359">Test the add-in</span></span>

1. <span data-ttu-id="7aea3-p148">Git bash ウィンドウまたは Node.JS 対応のシステム プロンプトが前の段階のチュートリアルから開いたままになっている場合は、Ctrl + C を 2 回入力して実行中の Web サーバーを停止します。 それ以外の場合は、Git bash ウィンドウまたは Node.JS 対応のシステム プロンプトを開いて、プロジェクトの **Start** フォルダーに移動します。</span><span class="sxs-lookup"><span data-stu-id="7aea3-p148">If the Git bash window, or Node.JS-enabled system prompt, from the previous stage tutorial is still open, enter Ctrl+C twice to stop the running web server. Otherwise, open a Git bash window, or Node.JS-enabled system prompt, and navigate to the **Start** folder of the project.</span></span>

     > [!NOTE]
     > <span data-ttu-id="7aea3-p149">ブラウザー同期サーバーは、app.js ファイルなどのファイルに変更を加えるたびに作業ウィンドウ内のアドインを再読み込みしますが、JavaScript を再トランスパイルしないため、ビルド コマンドを繰り返し実行して、app.js への変更を反映させる必要があります。 これを行うには、プロンプトが表示されてビルド コマンドを入力できるようにするため、サーバー プロセスを強制終了する必要があります。 ビルド後に、サーバーを再起動します。 次の数ステップで、このプロセスを実行します。</span><span class="sxs-lookup"><span data-stu-id="7aea3-p149">Although the browser-sync server reloads your add-in in the task pane every time you make a change to any file, including the app.js file, it does not retranspile the JavaScript, so you must repeat the build command in order for your changes to app.js to take effect. In order to do this, you need to kill the server process so that the prompt appears and you can enter the build command. After the build, restart the server. The next few steps carry out this process.</span></span>

2. <span data-ttu-id="7aea3-366">`npm run build` コマンドを実行し、Office アドインを実行できるすべてのホストでサポートされている以前のバージョンの JavaScript に ES6 ソース コードをトランスパイルします。</span><span class="sxs-lookup"><span data-stu-id="7aea3-366">Run the command `npm run build` to transpile your ES6 source code to an earlier version of JavaScript that is supported by all the hosts where Office Add-ins can run.</span></span>

3. <span data-ttu-id="7aea3-367">`npm start` コマンドを実行して、ローカルホストで稼働する Web サーバーを起動します。</span><span class="sxs-lookup"><span data-stu-id="7aea3-367">Run the command `npm start` to start a web server running on localhost.</span></span>

4. <span data-ttu-id="7aea3-368">作業ウィンドウを再読み込みするために、そのウィンドウを閉じて、**[ホーム]** メニューの **[作業ウィンドウの表示]** を選択してアドインを再度開きます。</span><span class="sxs-lookup"><span data-stu-id="7aea3-368">Reload the task pane by closing it, and then on the **Home** menu, select **Show Taskpane** to reopen the add-in.</span></span>

5. <span data-ttu-id="7aea3-369">作業ウィンドウで **[段落の挿入]** を少なくとも 3 回選択し、ドキュメントに段落がいくつかあることを確認します。</span><span class="sxs-lookup"><span data-stu-id="7aea3-369">In the task pane, choose **Insert Paragraph** at least three times to ensure that there are a few paragraphs in the document.</span></span>

6. <span data-ttu-id="7aea3-370">**[画像の挿入]** ボタンをクリックし、ドキュメントの末尾に画像が挿入されることに注意してください。</span><span class="sxs-lookup"><span data-stu-id="7aea3-370">Choose the **Insert Image** button and note that an image is inserted at the end of the document.</span></span>

7. <span data-ttu-id="7aea3-371">**[HTML の挿入]** ボタンをクリックし、ドキュメントの末尾に 2 つの段落が挿入され、最初の段落に Verdana フォントが設定されていることに注意してください。</span><span class="sxs-lookup"><span data-stu-id="7aea3-371">Choose the **Insert HTML** button and note that two paragraphs are inserted at the end of the document, and that the first one has Verdana font.</span></span>

8. <span data-ttu-id="7aea3-372">**[テーブルの挿入]** ボタンをクリックし、2 番目の段落の後にテーブルが挿入されることに注意してください。</span><span class="sxs-lookup"><span data-stu-id="7aea3-372">Choose the **Insert Table** button and note that a table is inserted after the second paragraph.</span></span>

    ![Word のチュートリアル - 画像、HTML、テーブルの挿入](../images/word-tutorial-insert-image-html-table.png)

## <a name="create-and-update-content-controls"></a><span data-ttu-id="7aea3-374">コンテンツ コントロールの作成と更新</span><span class="sxs-lookup"><span data-stu-id="7aea3-374">Create and update content controls</span></span>

<span data-ttu-id="7aea3-375">このチュートリアルの手順では、ドキュメント内にリッチ テキスト コンテンツ コントロールを作成する方法、およびそのコントロールにコンテンツを挿入したり置き換えたりする方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="7aea3-375">In this step of the tutorial, you'll learn how to create Rich Text content controls in the document, and then how to insert and replace content in the controls.</span></span>

> [!NOTE]
> <span data-ttu-id="7aea3-376">UI から Word 文書に追加できるコンテンツ コントロールにはいくつかの種類がありますが、Word.js では現在のところリッチ テキスト コンテンツ コントロールのみがサポートされています。</span><span class="sxs-lookup"><span data-stu-id="7aea3-376">There are several types of content controls that can be added to a Word document through the UI; but currently only Rich Text content controls are supported by Word.js.</span></span>
>
> <span data-ttu-id="7aea3-p150">チュートリアルのこの手順を開始する前に、Word UI からリッチ テキスト コンテンツ コントロールを作成して操作し、コントロールとそのプロパティを理解しておくことをお勧めします。 詳細については、「[ユーザーが Word 上で記入または印刷するフォームを作成する](https://support.office.com/article/create-forms-that-users-complete-or-print-in-word-040c5cc1-e309-445b-94ac-542f732c8c8b)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="7aea3-p150">Before you start this step of the tutorial, we recommend that you create and manipulate Rich Text content controls through the Word UI, so you can be familiar with the controls and their properties. For details, see [Create forms that users complete or print in Word](https://support.office.com/article/create-forms-that-users-complete-or-print-in-word-040c5cc1-e309-445b-94ac-542f732c8c8b).</span></span>

### <a name="create-a-content-control"></a><span data-ttu-id="7aea3-379">コンテンツ コントロールを作成する</span><span class="sxs-lookup"><span data-stu-id="7aea3-379">Create a content control</span></span>

1. <span data-ttu-id="7aea3-380">コード エディターでプロジェクトを開きます。</span><span class="sxs-lookup"><span data-stu-id="7aea3-380">Open the project in your code editor.</span></span>

2. <span data-ttu-id="7aea3-381">index.html ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="7aea3-381">Open the file index.html.</span></span>

3. <span data-ttu-id="7aea3-382">`replace-text` ボタンを格納している `div` の下に、次のマークアップを追加します。</span><span class="sxs-lookup"><span data-stu-id="7aea3-382">Below the `div` that contains the `replace-text` button, add the following markup:</span></span>

    ```html
    <div class="padding">
        <button class="ms-Button" id="create-content-control">Create Content Control</button>
    </div>
    ```

4. <span data-ttu-id="7aea3-383">app.js ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="7aea3-383">Open the app.js file.</span></span>

5. <span data-ttu-id="7aea3-384">`insert-table` ボタンにクリック ハンドラーを割り当てる行の下に、次のコードを追加します。</span><span class="sxs-lookup"><span data-stu-id="7aea3-384">Below the line that assigns a click handler to the `insert-table` button, add the following code:</span></span>

    ```js
    $('#create-content-control').click(createContentControl);
    ```

6. <span data-ttu-id="7aea3-385">`insertTable` 関数の下に、次の関数を追加します。</span><span class="sxs-lookup"><span data-stu-id="7aea3-385">Below the `insertTable` function, add the following function:</span></span>

    ```js
    function createContentControl() {
        Word.run(function (context) {

            // TODO1: Queue commands to create a content control.

            return context.sync();
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }
    ```

7. <span data-ttu-id="7aea3-386">`TODO1` を次のコードに置き換えます。</span><span class="sxs-lookup"><span data-stu-id="7aea3-386">Replace `TODO1` with the following code.</span></span> <span data-ttu-id="7aea3-387">次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="7aea3-387">Note:</span></span>

   - <span data-ttu-id="7aea3-p152">このコードの目的は、コンテンツ コントロール内の Office 365 という語句をラップすることです。 これは前提を単純化し、文字列は存在しており、ユーザーがその文字列を選択したものとしています。</span><span class="sxs-lookup"><span data-stu-id="7aea3-p152">This code is intended to wrap the phrase "Office 365" in a content control. It makes a simplifying assumption that the string is present and the user has selected it.</span></span>

   - <span data-ttu-id="7aea3-390">`ContentControl.title` プロパティは、コンテンツ コントロールの表示タイトルを指定します。</span><span class="sxs-lookup"><span data-stu-id="7aea3-390">The `ContentControl.title` property specifies the visible title of the content control.</span></span>

   - <span data-ttu-id="7aea3-391">`ContentControl.tag` プロパティは、`ContentControlCollection.getByTag` メソッドを使用してコンテンツ コントロールへの参照を取得するために使用できるタグを指定します。これを後述する関数で使用します。</span><span class="sxs-lookup"><span data-stu-id="7aea3-391">The `ContentControl.tag` property specifies an tag that can be used to get a reference to a content control using the `ContentControlCollection.getByTag` method, which you'll use in a later function.</span></span>

   - <span data-ttu-id="7aea3-p153">`ContentControl.appearance` プロパティは、コントロールの外観を指定します。 Tags という値を使用すると、コントロールは開始タグと終了タグにラップされます。開始タグには、コンテンツ コントロールのタイトルが設定されます。 その他の値として、BoundingBox と None が使用できます。</span><span class="sxs-lookup"><span data-stu-id="7aea3-p153">The `ContentControl.appearance` property specifies the visual look of the control. Using the value "Tags" means that the control will be wrapped in opening and closing tags, and the opening tag will have the content control's title. Other possible values are "BoundingBox" and "None".</span></span>

   - <span data-ttu-id="7aea3-395">`ContentControl.color` プロパティは、タグまたは境界ボックスの境界線の色を指定します。</span><span class="sxs-lookup"><span data-stu-id="7aea3-395">The `ContentControl.color` property specifies the color of the tags or the border of the bounding box.</span></span>

    ```js
    var serviceNameRange = context.document.getSelection();
    var serviceNameContentControl = serviceNameRange.insertContentControl();
    serviceNameContentControl.title = "Service Name";
    serviceNameContentControl.tag = "serviceName";
    serviceNameContentControl.appearance = "Tags";
    serviceNameContentControl.color = "blue";
    ```

### <a name="replace-the-content-of-the-content-control"></a><span data-ttu-id="7aea3-396">コンテンツ コントロールのコンテンツを置き換える</span><span class="sxs-lookup"><span data-stu-id="7aea3-396">Replace the content of the content control</span></span>

1. <span data-ttu-id="7aea3-397">index.html ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="7aea3-397">Open the file index.html.</span></span>

2. <span data-ttu-id="7aea3-398">`create-content-control` ボタンを格納している `div` の下に、次のマークアップを追加します。</span><span class="sxs-lookup"><span data-stu-id="7aea3-398">Below the `div` that contains the `create-content-control` button, add the following markup:</span></span>

    ```html
    <div class="padding">
        <button class="ms-Button" id="replace-content-in-control">Rename Service</button>
    </div>
    ```

3. <span data-ttu-id="7aea3-399">app.js ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="7aea3-399">Open the app.js file.</span></span>

4. <span data-ttu-id="7aea3-400">`create-content-control` ボタンにクリック ハンドラーを割り当てる行の下に、次のコードを追加します。</span><span class="sxs-lookup"><span data-stu-id="7aea3-400">Below the line that assigns a click handler to the `create-content-control` button, add the following code:</span></span>

    ```js
    $('#replace-content-in-control').click(replaceContentInControl);
    ```

5. <span data-ttu-id="7aea3-401">`createContentControl` 関数の下に、次の関数を追加します。</span><span class="sxs-lookup"><span data-stu-id="7aea3-401">Below the `createContentControl` function, add the following function:</span></span>

    ```js
    function replaceContentInControl() {
        Word.run(function (context) {

            // TODO1: Queue commands to replace the text in the Service Name
            //        content control.

            return context.sync();
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }
    ```

6. <span data-ttu-id="7aea3-p154">`TODO1` を次のコードに置き換えます。次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="7aea3-p154">Replace `TODO1` with the following code. Note:</span></span>

    - <span data-ttu-id="7aea3-404">`ContentControlCollection.getByTag` メソッドによって、指定されたタグのすべてのコンテンツ コントロールの `ContentControlCollection` が返されます。</span><span class="sxs-lookup"><span data-stu-id="7aea3-404">The `ContentControlCollection.getByTag` method returns a `ContentControlCollection` of all content controls of the specified tag.</span></span> <span data-ttu-id="7aea3-405">`getFirst` を使用して、目的のコントロールの参照を取得します。</span><span class="sxs-lookup"><span data-stu-id="7aea3-405">We use `getFirst` to get a reference to the desired control.</span></span>

    ```js
    var serviceNameContentControl = context.document.contentControls.getByTag("serviceName").getFirst();
    serviceNameContentControl.insertText("Fabrikam Online Productivity Suite", "Replace");
    ```

### <a name="test-the-add-in"></a><span data-ttu-id="7aea3-406">アドインをテストする</span><span class="sxs-lookup"><span data-stu-id="7aea3-406">Test the add-in</span></span>

1. <span data-ttu-id="7aea3-p156">Git bash ウィンドウまたは Node.JS 対応のシステム プロンプトが前の段階のチュートリアルから開いたままになっている場合は、Ctrl + C を 2 回入力して実行中の Web サーバーを停止します。 それ以外の場合は、Git bash ウィンドウまたは Node.JS 対応のシステム プロンプトを開いて、プロジェクトの **Start** フォルダーに移動します。</span><span class="sxs-lookup"><span data-stu-id="7aea3-p156">If the Git bash window, or Node.JS-enabled system prompt, from the previous stage tutorial is still open, enter Ctrl+C twice to stop the running web server. Otherwise, open a Git bash window, or Node.JS-enabled system prompt, and navigate to the **Start** folder of the project.</span></span>

     > [!NOTE]
     > <span data-ttu-id="7aea3-p157">ブラウザー同期サーバーは、app.js ファイルなどのファイルに変更を加えるたびに作業ウィンドウ内のアドインを再読み込みしますが、JavaScript を再トランスパイルしないため、ビルド コマンドを繰り返し実行して、app.js への変更を反映させる必要があります。 これを行うには、プロンプトが表示されてビルド コマンドを入力できるようにするため、サーバー プロセスを強制終了する必要があります。 ビルド後に、サーバーを再起動します。 次の数ステップで、このプロセスを実行します。</span><span class="sxs-lookup"><span data-stu-id="7aea3-p157">Although the browser-sync server reloads your add-in in the task pane every time you make a change to any file, including the app.js file, it does not retranspile the JavaScript, so you must repeat the build command in order for your changes to app.js to take effect. In order to do this, you need to kill the server process so that the prompt appears and you can enter the build command. After the build, restart the server. The next few steps carry out this process.</span></span>

2. <span data-ttu-id="7aea3-413">`npm run build` コマンドを実行し、Office アドインを実行できるすべてのホストでサポートされている以前のバージョンの JavaScript に ES6 ソース コードをトランスパイルします。</span><span class="sxs-lookup"><span data-stu-id="7aea3-413">Run the command `npm run build` to transpile your ES6 source code to an earlier version of JavaScript that is supported by all the hosts where Office Add-ins can run.</span></span>

3. <span data-ttu-id="7aea3-414">`npm start` コマンドを実行して、ローカルホストで稼働する Web サーバーを起動します。</span><span class="sxs-lookup"><span data-stu-id="7aea3-414">Run the command `npm start` to start a web server running on localhost.</span></span>

4. <span data-ttu-id="7aea3-415">作業ウィンドウを再読み込みするために、そのウィンドウを閉じて、**[ホーム]** メニューの **[作業ウィンドウの表示]** を選択してアドインを再度開きます。</span><span class="sxs-lookup"><span data-stu-id="7aea3-415">Reload the task pane by closing it, and then on the **Home** menu, select **Show Taskpane** to reopen the add-in.</span></span>

5. <span data-ttu-id="7aea3-416">作業ウィンドウで **[段落の挿入]** を選択し、文書の先頭が Office 365 となっている段落があることを確認します。</span><span class="sxs-lookup"><span data-stu-id="7aea3-416">In the task pane, choose **Insert Paragraph** to ensure that there is a paragraph with "Office 365" at the top of the document.</span></span>

6. <span data-ttu-id="7aea3-p158">追加した段落の Office 365 という語句を選択し、**[コンテンツ コントロールの作成]** ボタンを選択します。 Service Name というラベルが付いたタグで語句がラップされていることに注意してください。</span><span class="sxs-lookup"><span data-stu-id="7aea3-p158">Select the phrase "Office 365" in the paragraph you just added, and then choose the **Create Content Control** button. Note that the phrase is wrapped in tags labelled "Service Name".</span></span>

7. <span data-ttu-id="7aea3-419">**[サービス名の変更]** ボタンを選択し、コンテンツ コントロールのテキストが Fabrikam Online Productivity Suite に変わることに注意してください。</span><span class="sxs-lookup"><span data-stu-id="7aea3-419">Choose the **Rename Service** button and note that the text of the content control changes to "Fabrikam Online Productivity Suite".</span></span>

    ![Word のチュートリアル - コンテンツ コントロールの作成とテキストの変更](../images/word-tutorial-content-control.png)

## <a name="next-steps"></a><span data-ttu-id="7aea3-421">次の手順</span><span class="sxs-lookup"><span data-stu-id="7aea3-421">Next steps</span></span>

<span data-ttu-id="7aea3-422">このチュートリアルでは、テキスト、画像、Word 文書の他のコンテンツを挿入および置換する Word 作業ウィンドウ アドインを作成しました。</span><span class="sxs-lookup"><span data-stu-id="7aea3-422">In this tutorial, you've created a Word task pane add-in that inserts and replaces text, images, and other content in a Word document.</span></span> <span data-ttu-id="7aea3-423">Word アドインの構築に関する詳細については、次の記事にお進みください。</span><span class="sxs-lookup"><span data-stu-id="7aea3-423">To learn more about building Word add-ins, continue to the following article:</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="7aea3-424">Word アドインの概要</span><span class="sxs-lookup"><span data-stu-id="7aea3-424">Word add-ins overview</span></span>](../word/word-add-ins-programming-overview.md)
