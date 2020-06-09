---
title: ドキュメントやスプレッドシート内の領域へのバインド
description: バインドを使用して、ドキュメントまたはスプレッドシートの特定の領域または要素を識別子を介して一貫した方法でアクセスできるようにする方法について説明します。
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: 8fe754dedc2313f2a56dab71491dbdf8026085e1
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/08/2020
ms.locfileid: "44608391"
---
# <a name="bind-to-regions-in-a-document-or-spreadsheet"></a><span data-ttu-id="358ee-103">ドキュメントやスプレッドシート内の領域へのバインド</span><span class="sxs-lookup"><span data-stu-id="358ee-103">Bind to regions in a document or spreadsheet</span></span>

<span data-ttu-id="358ee-p101">バインドベースのデータ アクセスにより、コンテンツ アドインおよび作業ウィンドウ アドインは、ドキュメントまたはスプレッドシートの特定の領域に ID を通じて一貫性をもってアクセスできます。アドインは、最初に、ドキュメントの部分と一意の ID を関連付けるいずれかのメソッド ([addFromPromptAsync]、[addFromSelectionAsync]、または [addFromNamedItemAsync]) を呼び出すことによって、バインドを確立する必要があります。バインドが確立されると、アドインは提供された ID を使用して、ドキュメントまたはスプレッドシート内の関連付けられた領域に含まれるデータにアクセスできます。バインドの作成により、アドインに次の効果がもたらされます。</span><span class="sxs-lookup"><span data-stu-id="358ee-p101">Binding-based data access enables content and task pane add-ins to consistently access a particular region of a document or spreadsheet through an identifier. The add-in first needs to establish the binding by calling one of the methods that associates a portion of the document with a unique identifier: [addFromPromptAsync], [addFromSelectionAsync], or [addFromNamedItemAsync]. After the binding is established, the add-in can use the provided identifier to access the data contained in the associated region of the document or spreadsheet. Creating bindings provides the following value to your add-in:</span></span>


- <span data-ttu-id="358ee-108">表、範囲、またはテキスト (隣接する一連の文字) など、サポートされている Office アプリケーション全体に共通のデータ構造へのアクセスを許可します。</span><span class="sxs-lookup"><span data-stu-id="358ee-108">Permits access to common data structures across supported Office applications, such as: tables, ranges, or text (a contiguous run of characters).</span></span>

- <span data-ttu-id="358ee-109">ユーザーによる選択を必要とせずに、読み取り/書き込み操作ができます。</span><span class="sxs-lookup"><span data-stu-id="358ee-109">Enables read/write operations without requiring the user to make a selection.</span></span>

- <span data-ttu-id="358ee-p102">アドインとドキュメント内のデータの間にリレーションシップが確立されます。バインドはドキュメント内に保持され、後でアクセスできます。</span><span class="sxs-lookup"><span data-stu-id="358ee-p102">Establishes a relationship between the add-in and the data in the document. Bindings are persisted in the document, and can be accessed at a later time.</span></span>

<span data-ttu-id="358ee-p103">また、バインドを確立すると、ドキュメントまたはスプレッドシートの特定の領域を範囲とする、データおよび選択範囲の変更イベントをサブスクライブできます。つまり、ドキュメントまたはスプレッドシート全体の全般的な変更ではなく、バインドされた領域内で発生する変更のみがアドインに通知されます。</span><span class="sxs-lookup"><span data-stu-id="358ee-p103">Establishing a binding also allows you to subscribe to data and selection change events that are scoped to that particular region of the document or spreadsheet. This means that the add-in is only notified of changes that happen within the bound region as opposed to general changes across the whole document or spreadsheet.</span></span>

<span data-ttu-id="358ee-p104">[Bindings] オブジェクトの [getAllAsync] メソッドを使用すると、ドキュメントまたはスプレッドシートに設定されているすべてのバインドのセットにアクセスできます。個々のバインドには、Bindings.[getByIdAsync] メソッドまたは [Office.select] メソッドを使用して ID でアクセスできます。[Bindings] オブジェクトの [addFromSelectionAsync]、[addFromPromptAsync]、[addFromNamedItemAsync]、または [releaseByIdAsync] メソッドのいずれかを使用して、新しいバインドを設定したり、既存のバインドを削除したりできます。</span><span class="sxs-lookup"><span data-stu-id="358ee-p104">The [Bindings] object exposes a [getAllAsync] method that gives access to the set of all bindings established on the document or spreadsheet. An individual binding can be accessed by its ID using either the Bindings.[getByIdAsync] or [Office.select] methods. You can establish new bindings as well as remove existing ones by using one of the following methods of the [Bindings] object: [addFromSelectionAsync], [addFromPromptAsync], [addFromNamedItemAsync], or [releaseByIdAsync].</span></span>


## <a name="binding-types"></a><span data-ttu-id="358ee-117">バインドの種類</span><span class="sxs-lookup"><span data-stu-id="358ee-117">Binding types</span></span>

<span data-ttu-id="358ee-118">[Addfromselectionasync]、 [addFromPromptAsync] 、または[addfromnameditemasync]メソッドを使用してバインドを作成する場合、 _bindingtype_パラメーターで指定する[バインドの]種類は3種類あり[ます]。</span><span class="sxs-lookup"><span data-stu-id="358ee-118">There are [three different types of bindings][Office.BindingType] that you specify with the _bindingType_ parameter when you create a binding with the [addFromSelectionAsync], [addFromPromptAsync] or [addFromNamedItemAsync] methods:</span></span>

1. <span data-ttu-id="358ee-119">**[テキスト バインド][TextBinding]** - テキストとして表現できるドキュメントの領域にバインドします。</span><span class="sxs-lookup"><span data-stu-id="358ee-119">**[Text Binding][TextBinding]** - Binds to a region of the document that can be represented as text.</span></span>

    <span data-ttu-id="358ee-p105">Word では、連続する選択範囲の大部分が有効ですが、Excel では、単一セルの範囲のみがテキスト バインドの対象です。Excel では、プレーン テキストのみがサポートされます。Word では、3 つの形式 (プレーン テキスト、HTML、および Open XML for Office) がサポートされます。</span><span class="sxs-lookup"><span data-stu-id="358ee-p105">In Word, most contiguous selections are valid, while in Excel only single cell selections can be the target of a text binding. In Excel, only plain text is supported. In Word, three formats are supported: plain text, HTML, and Open XML for Office.</span></span>

2. <span data-ttu-id="358ee-123">**[マトリックスバインド][MatrixBinding] \*\* -ヘッダーのない表形式データが含まれるドキュメントの固定領域にバインドします。マトリックスバインド内のデータは、2次元**配列\*\*として書き込みまたは読み取りが行われます。 JavaScript では、配列の配列として実装されています。</span><span class="sxs-lookup"><span data-stu-id="358ee-123">**[Matrix Binding][MatrixBinding]** - Binds to a fixed region of a document that contains tabular data without headers.Data in a matrix binding is written or read as a two dimensional **Array**, which in JavaScript is implemented as an array of arrays.</span></span> <span data-ttu-id="358ee-124">たとえば、2つの列の**文字列**値の2行を書き込みまたは読み込むことができ、 ` [['a', 'b'], ['c', 'd']]` 3 つの行の1つの列を書き込みまたはとして読み取ることができ `[['a'], ['b'], ['c']]` ます。</span><span class="sxs-lookup"><span data-stu-id="358ee-124">For example, two rows of **string** values in two columns can be written or read as ` [['a', 'b'], ['c', 'd']]`, and a single column of three rows can be written or read as `[['a'], ['b'], ['c']]`.</span></span>

    <span data-ttu-id="358ee-p107">Excel では、セルの連続する選択範囲を使用してマトリックス バインドを設定できます。Word では、表のみがマトリックス バインドをサポートします。</span><span class="sxs-lookup"><span data-stu-id="358ee-p107">In Excel, any contiguous selection of cells can be used to establish a matrix binding. In Word, only tables support matrix binding.</span></span>

3. <span data-ttu-id="358ee-p108">**[テーブル バインド][TableBinding]** - ヘッダーがある表が含まれるドキュメントの領域にバインドします。テーブル バインド内のデータは、[TableData](/javascript/api/office/office.tabledata) オブジェクトとして書き込みまたは読み取りが行われます。`TableData` オブジェクトは `headers` および `rows` プロパティを通じてデータを公開します。</span><span class="sxs-lookup"><span data-stu-id="358ee-p108">**[Table Binding][TableBinding]** - Binds to a region of a document that contains a table with headers.Data in a table binding is written or read as a [TableData](/javascript/api/office/office.tabledata) object. The `TableData` object exposes the data through the `headers` and `rows` properties.</span></span>

    <span data-ttu-id="358ee-p109">Excel または Word の表はすべて、テーブル バインドの基礎にできます。テーブル バインドを確立すると、ユーザーが表に追加する新しい各行または各列が、自動的にバインドに含まれます。</span><span class="sxs-lookup"><span data-stu-id="358ee-p109">Any Excel or Word table can be the basis for a table binding. After you establish a table binding, each new row or column a user adds to the table is automatically included in the binding.</span></span>

<span data-ttu-id="358ee-131">オブジェクトの3つの "addFrom" メソッドのいずれかを使用してバインドを作成した後は、 `Bindings` 対応するオブジェクトのメソッドである[MatrixBinding]、 [tablebinding]、または[textbinding]を使用してバインドのデータおよびプロパティを操作できます。</span><span class="sxs-lookup"><span data-stu-id="358ee-131">After a binding is created by using one of the three "addFrom" methods of the `Bindings` object, you can work with the binding's data and properties by using the methods of the corresponding object: [MatrixBinding], [TableBinding], or [TextBinding].</span></span> <span data-ttu-id="358ee-132">この 3 つのオブジェクトはすべて、[] オブジェクトの [getDataAsync] メソッドおよび `Binding` メソッドを継承しているので、バインドされたデータを操作できます。</span><span class="sxs-lookup"><span data-stu-id="358ee-132">All three of these objects inherit the [getDataAsync] and [setDataAsync] methods of the `Binding` object that enable you to interact with the bound data.</span></span>

> [!NOTE]
> <span data-ttu-id="358ee-p111">**マトリックス バインドとテーブル バインドの使い分け**作業中の表形式のデータに集計行が含まれ、アドインのスクリプトが集計行の値にアクセスする必要がある場合、またはユーザーの選択が集計行にあることを検出する必要がある場合は、マトリックス バインドを使用する必要があります。集計行を含む表形式データに対するテーブル バインドを設定する場合、[TableBinding.rowCount] プロパティおよびイベント ハンドラーの [BindingSelectionChangedEventArgs] オブジェクトの `rowCount` および `startRow` プロパティは、集計行のそれらの値に反映されません。この制限を回避するには、集計行を処理するマトリックス バインドを設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="358ee-p111">**When should you use matrix versus table bindings?** When the tabular data you are working with contains a total row, you must use a matrix binding if your add-in's script needs to access values in the total row or detect that the user's selection is in the total row. If you establish a table binding for tabular data that contains a total row, the [TableBinding.rowCount] property and the `rowCount` and `startRow` properties of the [BindingSelectionChangedEventArgs] object in event handlers won't reflect the total row in their values. To work around this limitation, you must use establish a matrix binding to work with the total row.</span></span>

## <a name="add-a-binding-to-the-users-current-selection"></a><span data-ttu-id="358ee-137">ユーザーの現在の選択範囲にバインドを追加する</span><span class="sxs-lookup"><span data-stu-id="358ee-137">Add a binding to the user's current selection</span></span>

<span data-ttu-id="358ee-138">次の例は、[addFromSelectionAsync] メソッドを使用して、ドキュメントの現在の選択範囲に `myBinding` というテキスト バインドを追加する方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="358ee-138">The following example shows how to add a text binding called  `myBinding` to the current selection in a document by using the [addFromSelectionAsync] method.</span></span>


```js
Office.context.document.bindings.addFromSelectionAsync(Office.BindingType.Text, { id: 'myBinding' }, function (asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        write('Action failed. Error: ' + asyncResult.error.message);
    } else {
        write('Added new binding with type: ' + asyncResult.value.type + ' and id: ' + asyncResult.value.id);
    }
});

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

<span data-ttu-id="358ee-p112">この例では、指定したバインドの種類はテキストです。つまり、選択範囲に対して [TextBinding] が作成されます。バインドが備えているデータと操作はバインドの種類ごとに異なります。[Office.BindingType] は、使用できるバインドの種類の値を示す列挙型です。</span><span class="sxs-lookup"><span data-stu-id="358ee-p112">In this example, the specified binding type is text. This means that a [TextBinding] will be created for the selection. Different binding types expose different data and operations. [Office.BindingType] is an enumeration of available binding type values.</span></span>

<span data-ttu-id="358ee-p113">2 番目のオプションのパラメーターは、作成している新しいバインドの ID を指定するオブジェクトです。指定しない場合、ID は自動的に生成されます。</span><span class="sxs-lookup"><span data-stu-id="358ee-p113">The second optional parameter is an object that specifies the ID of the new binding being created. If an ID is not specified, one is generated automatically.</span></span>

<span data-ttu-id="358ee-p114">最後の _callback_ パラメーターで関数に渡される匿名関数は、バインドの作成が完了したときに実行されます。この関数の単一のパラメーター [ を通じて、呼び出しの状態を示す ]AsyncResult`asyncResult` オブジェクトにアクセスできます。`AsyncResult.value` プロパティには、新規作成するバインドとして指定した種類の [Binding] オブジェクトへの参照が格納されます。この [Binding] オブジェクトを使用して、データを取得および設定できます。</span><span class="sxs-lookup"><span data-stu-id="358ee-p114">The anonymous function that is passed into the function as the final  _callback_ parameter is executed when the creation of the binding is complete. The function is called with a single parameter, `asyncResult`, which provides access to an [AsyncResult] object that provides the status of the call. The `AsyncResult.value` property contains a reference to a [Binding] object of the type that is specified for the newly created binding. You can use this [Binding] object to get and set data.</span></span>

## <a name="add-a-binding-from-a-prompt"></a><span data-ttu-id="358ee-149">プロンプトからバインドを追加する</span><span class="sxs-lookup"><span data-stu-id="358ee-149">Add a binding from a prompt</span></span>

<span data-ttu-id="358ee-p115">次の例は、[addFromPromptAsync] メソッドを使用して `myBinding` という名前のテキスト バインドを追加する方法を示しています。このメソッドでは、ユーザーはアプリケーションの組み込み範囲選択プロンプトを使用してバインドの範囲を指定できます。</span><span class="sxs-lookup"><span data-stu-id="358ee-p115">The following example shows how to add a text binding called  `myBinding` by using the [addFromPromptAsync] method. This method lets the user specify the range for the binding by using the application's built-in range selection prompt.</span></span>


```js
function bindFromPrompt() {
    Office.context.document.bindings.addFromPromptAsync(Office.BindingType.Text, { id: 'myBinding' }, function (asyncResult) {
        if (asyncResult.status == Office.AsyncResultStatus.Failed) {
            write('Action failed. Error: ' + asyncResult.error.message);
        } else {
            write('Added new binding with type: ' + asyncResult.value.type + ' and id: ' + asyncResult.value.id);
        }
    });
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message;
}
```

<span data-ttu-id="358ee-p116">この例では、指定されているバインドの種類はテキストです。つまり、ユーザーがプロンプトで指定した選択範囲に対して [TextBinding] が作成されます。</span><span class="sxs-lookup"><span data-stu-id="358ee-p116">In this example, the specified binding type is text. This means that a [TextBinding] will be created for the selection that the user specifies in the prompt.</span></span>

<span data-ttu-id="358ee-p117">2 番目のパラメーターは、作成している新しいバインドの ID を含むオブジェクトです。指定しない場合、ID は自動的に生成されます。</span><span class="sxs-lookup"><span data-stu-id="358ee-p117">The second parameter is an object that contains the ID of the new binding being created. If an ID is not specified, one is generated automatically.</span></span>

<span data-ttu-id="358ee-p118">3 番目の _callback_ パラメーターとして関数に渡される匿名関数は、バインドの作成が完了すると実行されます。コールバック関数が実行されると、[AsyncResult] オブジェクトには呼び出しのステータスおよび新しく作成されたバインドが格納されます。</span><span class="sxs-lookup"><span data-stu-id="358ee-p118">The anonymous function passed into the function as the third  _callback_ parameter is executed when the creation of the binding is complete. When the callback function executes, the [AsyncResult] object contains the status of the call and the newly created binding.</span></span>

<span data-ttu-id="358ee-158">図 1 は、Excel の組み込み範囲選択プロンプトを示しています。</span><span class="sxs-lookup"><span data-stu-id="358ee-158">Figure 1 shows the built-in range selection prompt in Excel.</span></span>


<span data-ttu-id="358ee-159">*図 1.Excel のデータ選択 UI*</span><span class="sxs-lookup"><span data-stu-id="358ee-159">*Figure 1. Excel Select Data UI*</span></span>

![Excel のデータ選択 UI](../images/agave-api-overview-excel-selection-ui.png)


## <a name="add-a-binding-to-a-named-item"></a><span data-ttu-id="358ee-161">名前付きアイテムにバインドを追加する</span><span class="sxs-lookup"><span data-stu-id="358ee-161">Add a binding to a named item</span></span>


<span data-ttu-id="358ee-162">次の例は、 `myRange` [Addfromnameditemasync]メソッドを使用して、既存の名前付きアイテムへのバインドを "matrix" バインドとして追加し、そのバインドを `id` "mymatrix" として割り当てる方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="358ee-162">The following example shows how to add a binding to the existing `myRange` named item as a "matrix" binding by using the [addFromNamedItemAsync] method, and assigns the binding's `id` as "myMatrix".</span></span>


```js
function bindNamedItem() {
    Office.context.document.bindings.addFromNamedItemAsync("myRange", "matrix", {id:'myMatrix'}, function (result) {
        if (result.status == 'succeeded'){
            write('Added new binding with type: ' + result.value.type + ' and id: ' + result.value.id);
            }
        else
            write('Error: ' + result.error.message);
    });
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}

```

<span data-ttu-id="358ee-163">**Excel**では、 `itemName` [Addfromnameditemasync]メソッドのパラメーターは、既存の名前付き範囲、参照スタイルで指定された範囲、 `A1` `("A1:A3")` またはテーブルを参照できます。</span><span class="sxs-lookup"><span data-stu-id="358ee-163">**For Excel**, the `itemName` parameter of the [addFromNamedItemAsync] method can refer to an existing named range, a range specified with the `A1` reference style `("A1:A3")`, or a table.</span></span> <span data-ttu-id="358ee-164">既定では、Excel のテーブルを追加すると、最初に追加したテーブルには "Table1"、次に追加したテーブルには "Table2" という名前が割り当てられます。</span><span class="sxs-lookup"><span data-stu-id="358ee-164">By default, adding a table in Excel assigns the name "Table1" for the first table you add, "Table2" for the second table you add, and so on.</span></span> <span data-ttu-id="358ee-165">Excel UI でテーブルにわかりやすい名前を割り当てるには、 `Table Name` [表のツール] のプロパティを使用します。 **Table Tools | Design**リボンの [デザイン] タブ</span><span class="sxs-lookup"><span data-stu-id="358ee-165">To assign a meaningful name for a table in the Excel UI, use the `Table Name` property on the **Table Tools | Design** tab of the ribbon.</span></span>


> [!NOTE]
> <span data-ttu-id="358ee-166">Excel では、名前付きアイテムとしてテーブルを指定する場合、次の形式のテーブル名にワークシート名を含めるように、名前を完全修飾する必要があります。`"Sheet1!Table1"`</span><span class="sxs-lookup"><span data-stu-id="358ee-166">In Excel, when specifying a table as a named item, you must fully qualify the name to include the worksheet name in the name of the table in this format: `"Sheet1!Table1"`</span></span>

<span data-ttu-id="358ee-167">次の例では、Excel で、列 A () 内の最初の3つのセルにバインドを作成し、 `"A1:A3"` id を割り当て、 `"MyCities"` そのバインドに3つの都市名を書き込みます。</span><span class="sxs-lookup"><span data-stu-id="358ee-167">The following example creates a binding in Excel to the first three cells in column A ( `"A1:A3"`), assigns the id `"MyCities"`, and then writes three city names to that binding.</span></span>


```js
 function bindingFromA1Range() {
    Office.context.document.bindings.addFromNamedItemAsync("A1:A3", "matrix", {id: "MyCities" },
        function (asyncResult) {
            if (asyncResult.status == "failed") {
                write('Error: ' + asyncResult.error.message);
            }
            else {
                // Write data to the new binding.
                Office.select("bindings#MyCities").setDataAsync([['Berlin'], ['Munich'], ['Duisburg']], { coercionType: "matrix" },
                    function (asyncResult) {
                        if (asyncResult.status == "failed") {
                            write('Error: ' + asyncResult.error.message);
                        }
                    });
            }
        });
}
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

<span data-ttu-id="358ee-168">**Word の**場合、 `itemName` [Addfromnameditemasync]メソッドのパラメーターは、 `Title` コンテンツコントロールのプロパティを参照し `Rich Text` ます。</span><span class="sxs-lookup"><span data-stu-id="358ee-168">**For Word**, the `itemName` parameter of the [addFromNamedItemAsync] method refers to the `Title` property of a `Rich Text` content control.</span></span> <span data-ttu-id="358ee-169">(`Rich Text` コンテンツ コントロール以外のコンテンツ コントロールにはバインドできません)。</span><span class="sxs-lookup"><span data-stu-id="358ee-169">(You can't bind to content controls other than the `Rich Text` content control.)</span></span>

<span data-ttu-id="358ee-170">既定では、コンテンツコントロールには値が割り当てられていません `Title*` 。</span><span class="sxs-lookup"><span data-stu-id="358ee-170">By default, a content control has no `Title*`value assigned.</span></span> <span data-ttu-id="358ee-171">Word UI で意味のあるテーブル名を割り当てるには、リボンの [ **開発者**] タブの [ **コントロール**] グループから [ **リッチ テキスト**] コンテンツ コントロールを挿入した後、[ **コントロール**] グループの [ **プロパティ**] コマンドを使用して [ **コンテンツ コントロールのプロパティ**] ダイアログ ボックスを表示します。</span><span class="sxs-lookup"><span data-stu-id="358ee-171">To assign a meaningful name in the Word UI, after inserting a **Rich Text** content control from the **Controls** group on the **Developer** tab of the ribbon, use the **Properties** command in the **Controls** group to display the **Content Control Properties** dialog box.</span></span> <span data-ttu-id="358ee-172">次に、 `Title` コンテンツコントロールのプロパティを、コードから参照する名前に設定します。</span><span class="sxs-lookup"><span data-stu-id="358ee-172">Then set the `Title` property of the content control to the name you want to reference from your code.</span></span>

<span data-ttu-id="358ee-173">次の例では、という名前のリッチテキストコンテンツコントロールに対して Word のテキストバインドを作成し、 `"FirstName"` **id** `"firstName"` を割り当て、その情報を表示します。</span><span class="sxs-lookup"><span data-stu-id="358ee-173">The following example creates a text binding in Word to a rich text content control named `"FirstName"`, assigns the **id** `"firstName"`, and then displays that information.</span></span>


```js
function bindContentControl() {
    Office.context.document.bindings.addFromNamedItemAsync('FirstName', 
        Office.BindingType.Text, {id:'firstName'},
        function (result) {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                write('Control bound. Binding.id: '
                    + result.value.id + ' Binding.type: ' + result.value.type);
            } else {
                write('Error:', result.error.message);
            }
    });
}
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

## <a name="get-all-bindings"></a><span data-ttu-id="358ee-174">すべてのバインドを取得する</span><span class="sxs-lookup"><span data-stu-id="358ee-174">Get all bindings</span></span>


<span data-ttu-id="358ee-175">次の例は、Bindings.[getAllAsync] メソッドを使用して、ドキュメント内のすべてのバインドを取得する方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="358ee-175">The following example shows how to get all bindings in a document by using the Bindings.[getAllAsync] method.</span></span>


```js
Office.context.document.bindings.getAllAsync(function (asyncResult) {
    var bindingString = '';
    for (var i in asyncResult.value) {
        bindingString += asyncResult.value[i].id + '\n';
    }
    write('Existing bindings: ' + bindingString);
});

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

<span data-ttu-id="358ee-176">パラメーターとして関数に渡される匿名関数 `callback` は、操作が完了したときに実行されます。</span><span class="sxs-lookup"><span data-stu-id="358ee-176">The anonymous function that is passed into the function as the `callback` parameter is executed when the operation is complete.</span></span> <span data-ttu-id="358ee-177">関数は、ドキュメント内のバインドの配列を含む単一のパラメーターで呼び出され `asyncResult` ます。</span><span class="sxs-lookup"><span data-stu-id="358ee-177">The function is called with a single parameter, `asyncResult`, which contains an array of the bindings in the document.</span></span> <span data-ttu-id="358ee-178">配列は反復処理されて、バインドの ID を含む文字列が作成されます。</span><span class="sxs-lookup"><span data-stu-id="358ee-178">The array is iterated to build a string that contains the IDs of the bindings.</span></span> <span data-ttu-id="358ee-179">この文字列がメッセージ ボックスに表示されます。</span><span class="sxs-lookup"><span data-stu-id="358ee-179">The string is then displayed in a message box.</span></span>


## <a name="get-a-binding-by-id-using-the-getbyidasync-method-of-the-bindings-object"></a><span data-ttu-id="358ee-180">Bindings オブジェクトの getByIdAsync メソッドを使用して ID でバインドを取得する</span><span class="sxs-lookup"><span data-stu-id="358ee-180">Get a binding by ID using the getByIdAsync method of the Bindings object</span></span>


<span data-ttu-id="358ee-p123">次の例は、[getByIdAsync] メソッドを使用し、ID を指定してドキュメント内のバインドを取得する方法を示しています。この例では、前述のメソッドのいずれかを使用して `'myBinding'` という名前のバインドがドキュメントに追加されたと想定しています。</span><span class="sxs-lookup"><span data-stu-id="358ee-p123">The following example shows how to use the [getByIdAsync] method to get a binding in a document by specifying its ID. This example assumes that a binding named `'myBinding'` was added to the document using one of the methods described earlier in this topic.</span></span>


```js
Office.context.document.bindings.getByIdAsync('myBinding', function (asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        write('Action failed. Error: ' + asyncResult.error.message);
    } 
    else {
        write('Retrieved binding with type: ' + asyncResult.value.type + ' and id: ' + asyncResult.value.id);
    }
});

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

<span data-ttu-id="358ee-183">この例では、最初の `id` パラメーターは取得するバインドの ID です。</span><span class="sxs-lookup"><span data-stu-id="358ee-183">In the example, the first `id` parameter is the ID of the binding to retrieve.</span></span>

<span data-ttu-id="358ee-184">2番目の_callback_パラメーターとして関数に渡される匿名関数は、操作が完了したときに実行されます。</span><span class="sxs-lookup"><span data-stu-id="358ee-184">The anonymous function that is passed into the function as the second _callback_ parameter is executed when the operation is completed.</span></span> <span data-ttu-id="358ee-185">この関数は、呼び出しのステータスおよび ID が "myBinding" であるバインドが格納される _asyncResult_ という 1 つのパラメーターを使用して呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="358ee-185">The function is called with a single parameter, _asyncResult_, which contains the status of the call and the binding with the ID "myBinding".</span></span>


## <a name="get-a-binding-by-id-using-the-select-method-of-the-office-object"></a><span data-ttu-id="358ee-186">Office オブジェクトの select メソッドを使用して ID でバインドを取得する</span><span class="sxs-lookup"><span data-stu-id="358ee-186">Get a binding by ID using the select method of the Office object</span></span>


<span data-ttu-id="358ee-p125">次の例は、[Office.select] メソッドを使用してセレクター文字列に ID を指定することによって、ドキュメント内の [Binding] オブジェクトの promise を取得する方法を示しています。その後、Binding.[getDataAsync] メソッドを呼び出して、指定したバインドからデータを取得します。この例では、前述のメソッドのいずれかを使用して `'myBinding'` という名前のバインドがドキュメントに追加されたと想定しています。</span><span class="sxs-lookup"><span data-stu-id="358ee-p125">The following example shows how to use the [Office.select] method to get a [Binding] object promise in a document by specifying its ID in a selector string. It then calls the Binding.[getDataAsync] method to get data from the specified binding. This example assumes that a binding named `'myBinding'` was added to the document using one of the methods described earlier in this topic.</span></span>


```js
Office.select("bindings#myBinding", function onError(){}).getDataAsync(function (asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        write('Action failed. Error: ' + asyncResult.error.message);
    } else {
        write(asyncResult.value);
    }
});

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```


> [!NOTE]
> <span data-ttu-id="358ee-190">`select`メソッド promise が正常に[Binding]オブジェクトを返した場合、そのオブジェクトは、 [getdataasync]、 [setdataasync]、 [addハンドラ async]、および[removeハンドラ async]の4つのメソッドのみを公開します。</span><span class="sxs-lookup"><span data-stu-id="358ee-190">If the `select` method promise successfully returns a [Binding] object, that object exposes only the following four methods of the object: [getDataAsync], [setDataAsync], [addHandlerAsync], and [removeHandlerAsync].</span></span> <span data-ttu-id="358ee-191">Promise が Binding オブジェクトを返すことができない場合は、 `onError` コールバックを使用して[asyncResult]. error オブジェクトにアクセスし、さらに情報を取得することができます。メソッドによって返される[binding]オブジェクトの promise によって公開される4つのメソッド以外の binding オブジェクトのメンバーを呼び出す必要がある場合は `select` 、代わりに、[ドキュメントバインディング]のプロパティとバインドを使用して[getbyidasync]メソッドを使用します。[getByIdAsync][バインド]オブジェクトを取得する getByIdAsync メソッド。</span><span class="sxs-lookup"><span data-stu-id="358ee-191">If the promise cannot return a Binding object, the `onError` callback can be used to access an [asyncResult].error object to get more information.If you need to call a member of the Binding object other than the four methods exposed by the [Binding] object promise returned by the `select` method, instead use the [getByIdAsync] method by using the [Document.bindings] property and Bindings.[getByIdAsync] method to retrieve the [Binding] object.</span></span>

## <a name="release-a-binding-by-id"></a><span data-ttu-id="358ee-192">ID でバインドを解除する</span><span class="sxs-lookup"><span data-stu-id="358ee-192">Release a binding by ID</span></span>


<span data-ttu-id="358ee-193">次の例は、[releaseByIdAsync] メソッドを使用して ID を指定し、ドキュメント内のバインドを解除する方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="358ee-193">The following example shows how use the [releaseByIdAsync] method to release a binding in a document by specifying its ID.</span></span>

```js
Office.context.document.bindings.releaseByIdAsync('myBinding', function (asyncResult) {
    write('Released myBinding!');
});

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

<span data-ttu-id="358ee-194">この例で、最初の `id` パラメーターは解除するバインドの ID です。</span><span class="sxs-lookup"><span data-stu-id="358ee-194">In the example, the first `id` parameter is the ID of the binding to release.</span></span>

<span data-ttu-id="358ee-p127">2 番目のパラメーターとして関数に渡される匿名関数は、操作の完了時に実行されます。この関数は、呼び出しのステータスが格納される  [asyncResult] という 1 つのパラメーターを使用して呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="358ee-p127">The anonymous function that is passed into the function as the second parameter is a callback that is executed when the operation is complete. The function is called with a single parameter,  [asyncResult], which contains the status of the call.</span></span>


## <a name="read-data-from-a-binding"></a><span data-ttu-id="358ee-197">バインドからデータを読み取る</span><span class="sxs-lookup"><span data-stu-id="358ee-197">Read data from a binding</span></span>


<span data-ttu-id="358ee-198">次の例は、[getDataAsync] メソッドを使用して既存のバインドからデータを取得する方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="358ee-198">The following example shows how to use the [getDataAsync] method to get data from an existing binding.</span></span>


```js
myBinding.getDataAsync(function (asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        write('Action failed. Error: ' + asyncResult.error.message);
    } else {
        write(asyncResult.value);
    }
});

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

 <span data-ttu-id="358ee-p128">`myBinding` は、ドキュメント内の既存のテキスト バインドを格納している変数です。代わりに、[Office.select] メソッドを使用して ID によってバインドにアクセスし、次のように [getDataAsync] メソッドの呼び出しを開始できます。</span><span class="sxs-lookup"><span data-stu-id="358ee-p128">`myBinding` is a variable that contains an existing text binding in the document. Alternatively, you could use the [Office.select] to access the binding by its ID, and start your call to the [getDataAsync] method, like this:</span></span> 

```js 
Office.select("bindings#myBindingID").getDataAsync
```


<span data-ttu-id="358ee-p129">関数に渡される匿名関数は、操作の完了時に実行されるコールバックです。[AsyncResult].value プロパティには、`myBinding` 内のデータが格納されます。その値の型は、バインドの種類により異なります。この例のバインドはテキスト バインドです。そのため、値には文字列が格納されます。マトリックス バインドおよびテーブル バインドを使用して作業する追加の例については、[getDataAsync] メソッドのトピックを参照してください。</span><span class="sxs-lookup"><span data-stu-id="358ee-p129">The anonymous function that is passed into the function is a callback that is executed when the operation is complete. The [AsyncResult].value property contains the data within `myBinding`. The type of the value depends on the binding type. The binding in this example is a text binding. Therefore, the value will contain a string. For additional examples of working with matrix and table bindings, see the [getDataAsync] method topic.</span></span>


## <a name="write-data-to-a-binding"></a><span data-ttu-id="358ee-207">バインドにデータを書き込む</span><span class="sxs-lookup"><span data-stu-id="358ee-207">Write data to a binding</span></span>

<span data-ttu-id="358ee-208">次の例は、[setDataAsync] メソッドを使用して既存のバインドにデータを設定する方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="358ee-208">The following example shows how to use the [setDataAsync] method to set data in an existing binding.</span></span>

```js
myBinding.setDataAsync('Hello World!', function (asyncResult) { });
```

 <span data-ttu-id="358ee-209">`myBinding` は、ドキュメント内の既存のテキスト バインドを格納している変数です。</span><span class="sxs-lookup"><span data-stu-id="358ee-209">`myBinding` is a variable that contains an existing text binding in the document.</span></span>

<span data-ttu-id="358ee-210">この例では、最初のパラメーターは、に設定する値です `myBinding` 。</span><span class="sxs-lookup"><span data-stu-id="358ee-210">In the example, the first parameter is the value to set on `myBinding`.</span></span> <span data-ttu-id="358ee-211">これはテキスト バインドのため、値は `string` です。</span><span class="sxs-lookup"><span data-stu-id="358ee-211">Because this is a text binding, the value is a `string`.</span></span> <span data-ttu-id="358ee-212">バインドの種類が異なる場合、異なる型のデータが使用されます。</span><span class="sxs-lookup"><span data-stu-id="358ee-212">Different binding types accept different types of data.</span></span>

<span data-ttu-id="358ee-213">関数に渡される匿名関数は、操作の完了時に実行されるコールバックです。</span><span class="sxs-lookup"><span data-stu-id="358ee-213">The anonymous function that is passed into the function is a callback that is executed when the operation is complete.</span></span> <span data-ttu-id="358ee-214">この関数は、結果の状態を含む1つのパラメーターで呼び出され `asyncResult` ます。</span><span class="sxs-lookup"><span data-stu-id="358ee-214">The function is called with a single parameter, `asyncResult`, which contains the status of the result.</span></span>

> [!NOTE]
> <span data-ttu-id="358ee-215">Excel 2013 SP1 および Excel on the web の関連するビルドのリリースから、[バインド テーブルでデータの書き込みと更新を行う際に書式設定](../excel/excel-add-ins-tables.md)ができるようになりました。</span><span class="sxs-lookup"><span data-stu-id="358ee-215">Starting with the release of the Excel 2013 SP1 and the corresponding build of Excel on the web, you can now [set formatting when writing and updating data in bound tables](../excel/excel-add-ins-tables.md).</span></span>


## <a name="detect-changes-to-data-or-the-selection-in-a-binding"></a><span data-ttu-id="358ee-216">バインド内のデータまたは選択範囲の変更を検出する</span><span class="sxs-lookup"><span data-stu-id="358ee-216">Detect changes to data or the selection in a binding</span></span>


<span data-ttu-id="358ee-217">次の例は、ID が "MyBinding" であるバインドの [DataChanged](/javascript/api/office/office.binding) イベントにイベント ハンドラーを関連付ける方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="358ee-217">The following example shows how to attach an event handler to the [DataChanged](/javascript/api/office/office.binding) event of a binding with an id of "MyBinding".</span></span>


```js
function addHandler() {
Office.select("bindings#MyBinding").addHandlerAsync(
    Office.EventType.BindingDataChanged, dataChanged);
}
function dataChanged(eventArgs) {
    write('Bound data changed in binding: ' + eventArgs.binding.id);
}
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message;
}
```

<span data-ttu-id="358ee-218">`myBinding` は、ドキュメント内の既存のテキスト バインドを格納している変数です。</span><span class="sxs-lookup"><span data-stu-id="358ee-218">The `myBinding` is a variable that contains an existing text binding in the document.</span></span>

<span data-ttu-id="358ee-219">[Addハンドラ async]メソッドの最初の_eventType_パラメーターは、サブスクライブするイベントの名前を指定します。</span><span class="sxs-lookup"><span data-stu-id="358ee-219">The first _eventType_ parameter of the [addHandlerAsync] method specifies the name of the event to subscribe to.</span></span> <span data-ttu-id="358ee-220">[Office.EventType] は、使用できるイベントの種類の値の列挙型です。</span><span class="sxs-lookup"><span data-stu-id="358ee-220">[Office.EventType] is an enumeration of available event type values.</span></span> <span data-ttu-id="358ee-221">`Office.EventType.BindingDataChanged`文字列 "bindingDataChanged" に評価されます。</span><span class="sxs-lookup"><span data-stu-id="358ee-221">`Office.EventType.BindingDataChanged` evaluates to the string "bindingDataChanged".</span></span>

<span data-ttu-id="358ee-222">`dataChanged`2 番目の_handler_パラメーターとして関数に渡される関数は、バインド内のデータが変更されたときに実行されるイベントハンドラーです。</span><span class="sxs-lookup"><span data-stu-id="358ee-222">The `dataChanged` function that is passed into the function as the second _handler_ parameter is an event handler that is executed when the data in the binding is changed.</span></span> <span data-ttu-id="358ee-223">この関数は、バインドへの参照が格納される _eventArgs_ という 1 つのパラメーターを使用して呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="358ee-223">The function is called with a single parameter, _eventArgs_, which contains a reference to the binding.</span></span> <span data-ttu-id="358ee-224">このバインドを使用して、更新されたデータを取得できます。</span><span class="sxs-lookup"><span data-stu-id="358ee-224">This binding can be used to retrieve the updated data.</span></span>

<span data-ttu-id="358ee-p134">同様に、バインドの [SelectionChanged] イベントにイベント ハンドラーを関連付けることによって、バインド内の選択範囲の変更を検出できます。これを行うには、[addHandlerAsync] メソッドの `eventType` パラメーターを `Office.EventType.BindingSelectionChanged` または `"bindingSelectionChanged"` と指定します。</span><span class="sxs-lookup"><span data-stu-id="358ee-p134">Similarly, you can detect when a user changes selection in a binding by attaching an event handler to the [SelectionChanged] event of a binding. To do that, specify the `eventType` parameter of the [addHandlerAsync] method as `Office.EventType.BindingSelectionChanged` or `"bindingSelectionChanged"`.</span></span>

<span data-ttu-id="358ee-p135">[addHandlerAsync] メソッドを再び呼び出して、`handler` パラメーターに追加のイベント ハンドラー関数を指定すると、特定のイベントに複数のイベント ハンドラーを追加できます。この場合、各イベント ハンドラー関数の名前は一意である必要があります。</span><span class="sxs-lookup"><span data-stu-id="358ee-p135">You can add multiple event handlers for a given event by calling the [addHandlerAsync] method again and passing in an additional event handler function for the `handler` parameter. This will work correctly as long as the name of each event handler function is unique.</span></span>


### <a name="remove-an-event-handler"></a><span data-ttu-id="358ee-229">イベント ハンドラーを削除する</span><span class="sxs-lookup"><span data-stu-id="358ee-229">Remove an event handler</span></span>


<span data-ttu-id="358ee-p136">イベントのイベント ハンドラーを削除するには、最初の _eventType_ パラメーターにイベントの種類を指定し、2 番目の _handler_ パラメーターに削除するイベント ハンドラー関数の名前を指定して、[removeHandlerAsync] メソッドを呼び出します。たとえば、次の例では、前のセクションの例で追加した `dataChanged` イベント ハンドラー関数が削除されます。</span><span class="sxs-lookup"><span data-stu-id="358ee-p136">To remove an event handler for an event, call the [removeHandlerAsync] method passing in the event type as the first _eventType_ parameter, and the name of the event handler function to remove as the second _handler_ parameter. For example, the following function will remove the `dataChanged` event handler function added in the previous section's example.</span></span>


```js
function removeEventHandlerFromBinding() {
    Office.select("bindings#MyBinding").removeHandlerAsync(
        Office.EventType.BindingDataChanged, {handler:dataChanged});
}
```


> [!IMPORTANT]
> <span data-ttu-id="358ee-232">[Removehandler async]メソッドが呼び出されたときにオプションの_handler_パラメーターが省略された場合は、指定されたのすべてのイベントハンドラーが `eventType` 削除されます。</span><span class="sxs-lookup"><span data-stu-id="358ee-232">If the optional _handler_ parameter is omitted when the [removeHandlerAsync] method is called, all event handlers for the specified `eventType` will be removed.</span></span>


## <a name="see-also"></a><span data-ttu-id="358ee-233">関連項目</span><span class="sxs-lookup"><span data-stu-id="358ee-233">See also</span></span>

- [<span data-ttu-id="358ee-234">Office JavaScript API について</span><span class="sxs-lookup"><span data-stu-id="358ee-234">Understanding the Office JavaScript API</span></span>](understanding-the-javascript-api-for-office.md) 
- [<span data-ttu-id="358ee-235">Office アドインにおける非同期プログラミング</span><span class="sxs-lookup"><span data-stu-id="358ee-235">Asynchronous programming in Office Add-ins</span></span>](asynchronous-programming-in-office-add-ins.md)
- [<span data-ttu-id="358ee-236">ドキュメントやスプレッドシート内のアクティブな選択範囲へのデータの読み取りと書き込みを行います</span><span class="sxs-lookup"><span data-stu-id="358ee-236">Read and write data to the active selection in a document or spreadsheet</span></span>](read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md)

[Binding]:               /javascript/api/office/office.binding
[MatrixBinding]:         /javascript/api/office/office.matrixbinding
[TableBinding]:          /javascript/api/office/office.tablebinding
[TextBinding]:           /javascript/api/office/office.textbinding
[getDataAsync]:          /javascript/api/office/Office.Binding#getdataasync-options--callback-
[setDataAsync]:          /javascript/api/office/Office.Binding#setdataasync-data--options--callback-
[SelectionChanged]:      /javascript/api/office/office.bindingselectionchangedeventargs
[addHandlerAsync]:       /javascript/api/office/Office.Binding#addhandlerasync-eventtype--handler--options--callback-
[removeHandlerAsync]:    /javascript/api/office/Office.Binding#removehandlerasync-eventtype--options--callback-

[Bindings]:              /javascript/api/office/office.bindings
[getByIdAsync]:          /javascript/api/office/office.bindings#getbyidasync-id--options--callback- 
[getAllAsync]:           /javascript/api/office/office.bindings#getallasync-options--callback-
[addFromNamedItemAsync]: /javascript/api/office/office.bindings#addfromnameditemasync-itemname--bindingtype--options--callback-
[addFromSelectionAsync]: /javascript/api/office/office.bindings#addfromselectionasync-bindingtype--options--callback-
[addFromPromptAsync]:    /javascript/api/office/office.bindings#addfrompromptasync-bindingtype--options--callback-
[releaseByIdAsync]:      /javascript/api/office/office.bindings#releasebyidasync-id--options--callback-

[AsyncResult]:          /javascript/api/office/office.asyncresult
[Office.BindingType]:   /javascript/api/office/office.bindingtype
[Office.select]:        /javascript/api/office 
[Office.EventType]:     /javascript/api/office/office.eventtype 
[Document.bindings]:    /javascript/api/office/office.document


[TableBinding.rowCount]: /javascript/api/office/office.tablebinding
[BindingSelectionChangedEventArgs]: /javascript/api/office/office.bindingselectionchangedeventargs
