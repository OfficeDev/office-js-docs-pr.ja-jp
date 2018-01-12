
# <a name="office-javascript-api-support-for-content-and-task-pane-add-ins-in-office-2013"></a>Office 2013 でのコンテンツ アドインと作業ウィンドウ アドインの Office JavaScript API のサポート


[Office JavaScript API](../../reference/javascript-api-for-office.md) を使用して、Office 2013 ホスト アプリケーションの作業ウィンドウやコンテンツ用のアドインを作成できます。コンテンツと作業ウィンドウのアドインをサポートするオブジェクトとメソッドは、次のように分類されます。


1. **他の Office アドインと共有する共通のオブジェクト。** これらのオブジェクトには、[Office](../../reference/shared/office.md)、[Context](../../reference/shared/office.context.md)、および [AsyncResult](../../reference/shared/asyncresult.md) があります。**Office** オブジェクトは Office JavaScript API のルート オブジェクトです。**Context** オブジェクトはアドインのランタイム環境を表します。**Office** と **Context** は、いずれも Office アドインの基礎となるオブジェクトです。**AsyncResult** オブジェクトは、ユーザーが文書内で選択したものを読み取る **getSelectedDataAsync** メソッドに返されたデータなどの非同期操作の結果を表します。
    
2.  **Document オブジェクト。** コンテンツと作業ウィンドウのアドインに使用可能な API のほとんどは、[Document](../../reference/shared/document.md) オブジェクトのメソッド、プロパティ、およびイベントを通じて公開されます。コンテンツ アドインと作業ウィンドウ アドインは、[Office.context.document](../../reference/shared/office.context.document.md) プロパティを使用して **Document** オブジェクトにアクセスし、それを通して、[Bindings](../../reference/shared/bindings.bindings.md) オブジェクト、[CustomXmlParts](../../reference/shared/customxmlparts.customxmlparts.md) オブジェクト、[getSelectedDataAsync](../../reference/shared/document.getselecteddataasync.md) メソッド、[setSelectedDataAsync](../../reference/shared/document.setselecteddataasync.md) メソッド、[getFileAsync](../../reference/shared/document.getfileasync.md) メソッドなどの文書内のデータを操作するための API の主要メンバーにアクセスできます。**Document** オブジェクトは、文書が読み取り専用モードと編集モードのどちらになっているかを判断するための [mode](../../reference/shared/document.mode.md) プロパティと、現在の文書の URL を取得して、[Settings](../../reference/shared/document.url.md) オブジェクトにアクセスするための [url](../../reference/shared/settings.md) プロパティも提供します。**Document** オブジェクトは、[SelectionChanged](../../reference/shared/document.selectionchanged.event.md) イベントのイベント ハンドラの追加もサポートしているため、ユーザーが文書内の選択を変更した時点を検出できます。
    
   コンテンツ アドインや作業ウィンドウ アドインが **Document** オブジェクトにアクセスできるのは、DOM とランタイム環境が [Office.initialize](../../reference/shared/office.initialize.md) イベント用のイベント ハンドラーなどで読み込まれた後だけです。アドインが初期化されるときのイベント フローと、DOM とラインタイムが正常に読み込まれたかどうかの確認方法については、「[DOM とランタイム環境の読み込み](../../docs/develop/loading-the-dom-and-runtime-environment.md)」を参照してください。
    
3.  **特定の機能を操作するためのオブジェクト。**API の特定の機能を操作するには、次のオブジェクトとメソッドを使用します。
    
    - [Bindings](../../reference/shared/bindings.bindings.md) オブジェクトのメソッドを使用して、バインドを作成または取得します。また、[Binding](../../reference/shared/binding.md) オブジェクトのメソッドとプロパティを使用して、データを操作します。
    
    - [CustomXmlParts](../../reference/shared/customxmlparts.customxmlparts.md)、[CustomXmlPart](../../reference/shared/customxmlpart.customxmlpart.md)、および関連するオブジェクトを使用して、Word 文書内のカスタム XML パーツを作成および操作します。
    
    - [File](../../reference/shared/file.md) オブジェクトと [Slice](../../reference/shared/slice.md) オブジェクトを使用して、文書全体のコピーを作成し、それをチャンクまたは「スライス」に分割してから、それらのスライスに含まれるデータを読み取りまたは転送します。
    
    - [Settings](../../reference/shared/settings.md) オブジェクトを使用して、ユーザー設定やアドインの状態などのカスタム データを保存します。
    

 >**重要**  API メンバーの一部は、コンテンツ アドインと作業ウィンドウ アドインをホスト可能なすべての Office アプリケーションでサポートされているわけではありません。サポートされているメンバーを特定するには、次のいずれかを参照してください。

Office ホスト アプリケーション全体に渡る Office JavaScript API サポートの概要については、「[JavaScript API for Office について](../../docs/develop/understanding-the-javascript-api-for-office.md)」を参照してください。


## <a name="reading-and-writing-to-an-active-selection"></a>アクティブな選択範囲の読み取りと書き込み

文書、スプレッドシート、またはプレゼンテーション内のユーザーの現在の選択範囲に対して読み書きをすることができます。アドイン用のホスト アプリケーションに応じて、[Document](../../reference/shared/document.getselecteddataasync.md) オブジェクトの [getSelectedDataAsync](../../reference/shared/document.setselecteddataasync.md) メソッドと [setSelectedDataAsync](../../reference/shared/document.md) メソッド内のパラメーターとして読み書きするデータ構造のタイプを指定できます。たとえば、Word には任意のデータ タイプ (テキスト、HTML、表形式データ、または Office Open XML)、Excel にはテキストと表形式データ、および PowerPoint と Project にはテキストを指定できます。ユーザーの選択範囲に対する変更を検出するためのイベント ハンドラーを作成することもできます。以下の例では、**getSelectedDataAsync** メソッドを使用して、選択範囲からデータをテキストとして取得します。


```js
Office.context.document.getSelectedDataAsync(
    Office.CoercionType.Text, function (asyncResult) {
        if (asyncResult.status == Office.AsyncResultStatus.Failed) {
            write('Action failed. Error: ' + asyncResult.error.message);
        }
        else {
            write('Selected data: ' + asyncResult.value);
        }
    });

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}

```

詳細と例については、「[文書またはスプレッドシート内のアクティブな選択範囲へのデータの読み取りと書き込み](../../docs/develop/read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md)」を参照してください。


## <a name="binding-to-a-region-in-a-document-or-spreadsheet"></a>文書またはスプレッドシート内の領域へのバインド

**getSelectedDataAsync** メソッドと **setSelectedDataAsync** メソッドを使用して、文書、スプレッドシート、またはプレゼンテーション内のユーザーの*現在*の選択範囲を読み取りまたは書き込みできます。ただし、ユーザーに選択を要求せずにアドインの複数の実行セッションに渡って文書内の同じ領域にアクセスする場合は、最初にその領域をバインドする必要があります。そのバインドした領域に対するデータおよび選択範囲変更イベントにサブスクライブすることもできます。

バインドは、[Bindings](../../reference/shared/bindings.addfromnameditemasync.md) オブジェクトの [addFromNamedItemAsync](../../reference/shared/bindings.addfrompromptasync.md) メソッド、[addFromPromptAsync](../../reference/shared/bindings.addfromselectionasync.md) メソッド、または [addFromSelectionAsync](../../reference/shared/bindings.bindings.md) メソッドを使用して追加できます。これらのメソッドは、バインド内のデータにアクセスするため、あるいは、データ変更または選択範囲変更イベントにサブスクライブするために使用可能な識別子を返します。

**Bindings.addFromSelectionAsync** メソッドを使用して、文書内で現在選択されているテキストにバインドを追加する例を次に示します。



```js
Office.context.document.bindings.addFromSelectionAsync(
    Office.BindingType.Text, { id: 'myBinding' }, function (asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        write('Action failed. Error: ' + asyncResult.error.message);
    } else {
        write('Added new binding with type: ' +
            asyncResult.value.type + ' and id: ' + asyncResult.value.id);
    }
});

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

詳細と例については、「[文書またはスプレッドシート内の領域へのバインド](../../docs/develop/bind-to-regions-in-a-document-or-spreadsheet.md)」を参照してください。


## <a name="getting-entire-documents"></a>文書全体の取得

作業ウィンドウ アドインが PowerPoint または Word で実行される場合は、[Document.getFileAsync](../../reference/shared/document.getfileasync.md) メソッド、[File.getSliceAsync](../../reference/shared/file.getsliceasync.md) メソッド、および [File.closeAsync](../../reference/shared/file.closeasync.md) メソッドを使用して、プレゼンテーションまたは文書全体を取得できます。

**Document.getFileAsync** を呼び出すと、[File](../../reference/shared/file.md) オブジェクトに文書のコピーが返されます。**File** オブジェクトは、[Slice](../../reference/shared/document.md) オブジェクトとして表現される「チャンク」内の文書へのアクセスを提供します。**getFileAsync** を呼び出すときに、ファイル タイプ (テキストまたは圧縮 Open Office XML 形式) とスライスのサイズ (4 MB 以下) を指定できます。**File** オブジェクトの内容にアクセスするには、**Slice.data** プロパティに生データを返す [File.getSliceAsync](../../reference/shared/slice.data.md) を呼び出します。圧縮形式を指定した場合は、ファイル データがバイト配列で返されます。ファイルを Web サービスに転送する場合は、圧縮生データを base64 エンコード文字列に変換してから送信できます。最後に、ファイルのスライスの取得が完了したら、**File.closeAsync** メソッドを使用して文書を閉じます。

詳細については、「[PowerPoint または Word 用アドインからドキュメント全体を取得する方法](../../docs/develop/get-the-whole-document-from-an-add-in-for-powerpoint-or-word.md)」を参照してください。 


## <a name="reading-and-writing-custom-xml-parts-of-a-word-document"></a>Word 文書のカスタム XML パーツの読み取りと書き込み

Open Office XML ファイル形式とコンテンツ コントロールを使用すれば、Word 文書にカスタム XML パーツを追加して、その文書内のコンテンツ コントロールに XML パーツ内の要素をバインドすることができます。文書を開くと、Word がバインドされたコンテンツ コントロールを読み取り、カスタム XML パーツからのデータを自動的に設定します。ユーザーは、コンテンツ コントロールにデータを書き込むこともできます。ユーザーが文書を保存すると、コントロール内のデータがバインドされた XML パーツに保存されます。Word 用の作業ウィンドウ アドインは、[Document.customXmlParts](../../reference/shared/document.customxmlparts.md) プロパティ、[CustomXmlParts](../../reference/shared/customxmlparts.customxmlparts.md) オブジェクト、[CustomXmlPart](../../reference/shared/customxmlpart.customxmlpart.md) オブジェクト、および [CustomXmlNode](../../reference/shared/customxmlnode.customxmlnode.md) オブジェクトを使用して、文書に対して動的にデータを読み書きすることができます。

カスタム XML パーツは名前空間に関連付けることができます。名前空間内のカスタム XML パーツからデータを取得するには、[CustomXmlParts.getByNamespaceAsync](../../reference/shared/customxmlparts.getbynamespaceasync.md) メソッドを使用します。

[CustomXmlParts.getByIdAsync](../../reference/shared/customxmlparts.getbyidasync.md) メソッドを使用して、GUID でカスタム XML パーツにアクセスすることもできます。カスタム XML パーツを取得したら、[CustomXmlPart.getXmlAsync](../../reference/shared/customxmlpart.getxmlasync.md) メソッドを使用して XML データを取得します。

文書に新しいカスタム XML パーツを追加するには、**Document.customXmlParts** プロパティを使用して文書内に存在するカスタム XML パーツを取得し、[CustomXmlParts.addAsync](../../reference/shared/customxmlparts.addasync.md) メソッドを呼び出します。

作業ウィンドウ アドインでのカスタム XML パーツの操作方法の詳細については、「[Office Open XML を使用してより良い Word 用アドインを作成する](../../docs/word/create-better-add-ins-for-word-with-office-open-xml.md)」を参照してください。


## <a name="persisting-add-in-settings"></a>アドイン設定を保存する


多くの場合、ユーザー設定やアドインの状態など、アドインのカスタム データを保存し、次回、アドインを開いたとき、そのデータにアクセスする必要があります。一般的な Web プログラミング手法を利用し、ブラウザーの Cookie や HTML 5 Web ストレージなど、そのデータを保存できます。あるいは、アドインを Excel、PowerPoint、Word で実行する場合、[Settings](../../reference/shared/settings.md) オブジェクトのメソッドを使用できます。**Settings** オブジェクトで作成したデータは、アドインを挿入して保存したスプレッドシート、プレゼンテーション、文書に保存されます。このデータは、それを作成したアドインでのみ利用できます。

文書が保存されているサーバーとのやり取りを避けるために、**Settings** オブジェクトを使用して作成されたデータはランタイムでメモリ上で管理されます。過去に保存した設定データがアドインの初期化時にメモリに読み込まれ、そのデータに対する変更は [Settings.saveAsync](../../reference/shared/settings.saveasync.md) メソッドを呼び出したときにのみ文書に保存されます。内部的に、データはシリアル化された JSON オブジェクト内に名前と値のペアとして保存されます。データのメモリ内コピーに対してアイテムの読み取り、書き込み、および削除を実行するには、[Settings](../../reference/shared/settings.get.md) オブジェクトの [get](../../reference/shared/settings.set.md) メソッド、[set](../../reference/shared/settings.removehandlerasync.md) メソッド、および **remove** メソッドを使用します。次のコード行は、`themeColor` という名前の設定を作成して、その値を 'green' に設定する方法を示しています。




```js
Office.context.document.settings.set('themeColor', 'green');
```

**set** メソッドと **remove** メソッドを使用した設定データの作成または削除はそのデータのメモリ内コピーに対して機能するため、**saveAsync** を呼び出して、設定データに対する変更をアドインが操作する文書内に保持する必要があります。

**Settings** オブジェクトのメソッドを使用したカスタム データの操作方法の詳細については、「[アドインの状態と設定を保存する](../../docs/develop/persisting-add-in-state-and-settings.md)」を参照してください。


## <a name="reading-properties-of-a-project-document"></a>プロジェクト文書のプロパティの読み取り

作業ウィンドウ アドインが Project で動作する場合は、そのアドインでアクティブ プロジェクト内のプロジェクト フィールド、リソース、およびタスク フィールドの一部からデータを読み取ることができます。これを実現するには、追加の Project 固有の機能を提供するように [Document](../../reference/shared/projectdocument.projectdocument.md) オブジェクトを拡張する **ProjectDocument** オブジェクトのメソッドとイベントを使用します。

Project のデータの読み取り操作の例については、「[テキスト エディターを使用して Project 2013 用の作業ウィンドウ アドインを初めて作成する](../../docs/project/create-your-first-task-pane-add-in-for-project-by-using-a-text-editor.md)」を参照してください。


## <a name="permissions-model-and-governance"></a>アクセス許可モデルとガバナンス

アドインは、そのマニフェスト内の **Permissions** 要素を使用して、必要な機能のレベルにアクセスするためのアクセス許可を Office JavaScript API に要求します。たとえば、アドインで文書に対する読み取りと書き込みアクセス権が必要な場合は、そのマニフェストで `ReadWriteDocument` をその **Permissions** 要素内のテキスト値として指定する必要があります。アクセス許可はユーザーのプライバシーとセキュリティを保護するために存在しているので、ベスト プラクティスとしては、その機能に必要な最低限のアクセス許可を要求することをお勧めします。次の例は、作業ウィンドウのマニフェストで **ReadDocument** アクセス許可を要求する方法を示しています。


```XML
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.0"
 xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
 xsi:type="TaskPaneApp">
???<!-- Other manifest elements omitted. -->
  <Permissions>ReadDocument</Permissions>
???
</OfficeApp>

```

詳細については、「[コンテンツ アドインおよび作業ウィンドウ アドインでの API 使用のアクセス許可を要求する](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)」を参照してください。


## <a name="additional-resources"></a>その他の技術情報


- [Office の JavaScript API](../../reference/javascript-api-for-office.md)
    
- 
  [Office アドインのマニフェスト向けのスキーマ リファレンス](http://msdn.microsoft.com/en-us/library/7e0cadc3-f613-8eb9-57ef-9032cbb97f92.aspx)
    
- [Office アドインでのユーザー エラーのトラブルシューティング](../../docs/testing/testing-and-troubleshooting.md)
    
