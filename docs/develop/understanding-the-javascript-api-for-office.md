
# <a name="understanding-the-javascript-api-for-office"></a>JavaScript API for Office について



この記事では、JavaScript API for Office とその使用方法に関する情報を提供します。参照情報については、「[JavaScript API for Office](http://dev.office.com/reference/add-ins/javascript-api-for-office)」を参照してください。Visual Studio プロジェクト ファイルを JavaScript API for Office の最新バージョンに更新する方法については、「[JavaScript API for Office およびマニフェスト スキーマ ファイルのバージョンを更新する](../develop/update-your-javascript-api-for-office-and-manifest-schema-version.md)」を参照してください。

>
  **注:**アドインをビルドするとき、アドインを Office ストアに[発行](../publish/publish.md)する予定であれば、[Office ストア検証ポリシー](https://msdn.microsoft.com/en-us/library/jj220035.aspx)に準拠していることを確認してください。たとえば、検証に合格するには、アドインは、定義したメソッドをサポートするすべてのプラットフォーム全体で機能する必要があります (詳細については、[セクション 4.12](https://msdn.microsoft.com/en-us/library/jj220035.aspx#Anchor_3) と「[Office アドインを使用できるホストおよびプラットフォーム](https://dev.office.com/add-in-availability)」のページを参照してください)。

## <a name="referencing-the-javascript-api-for-office-library-in-your-add-in"></a>アドインで JavaScript API for Office ライブラリを参照する

[JavaScript API for Office](http://dev.office.com/reference/add-ins/javascript-api-for-office) ライブラリは、Office.js ファイルと関連するホスト アプリケーション固有のファイル (Excel-15.js や Outlook-15.js など) で構成されています。最も簡単に API を参照する方法は、次に示す `<script>` をページの `<head>` タグに追加して、CDN を使用することです。  

```html
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
```

これにより、アドインが最初に読み込まれるときに JavaScript API for Office ファイルのダウンロードとキャッシュを実行して、アドインが確実に指定したバージョンの最新の Office.js および関連ファイルを使用するようにします。

バージョン管理や下位互換性の処理方法など、Office.js CDN に関する詳細については、「[Office ライブラリの JavaScript API をそのコンテンツ配信ネットワーク (CDN) から参照する](referencing-the-javascript-api-for-office-library-from-its-cdn.md)」を参照してください。

## <a name="initializing-your-add-in"></a>アドインの初期化


 **適用対象:** すべてのアドインの種類


Office.js は、API が完全に読み込まれていてユーザーによる操作ができる状態になっているときに起動されたとしても初期化を提供します。**initialize** イベント ハンドラーを使用すると、ユーザーに Excel のセルを複数選択するように求めるメッセージを表示し、選択された値で初期化したグラフを挿入するなど、アドインの一般的な初期化シナリオを実装できるようになります。また、アドインのその他のカスタム ロジックを初期化する場合 (バインドを確立する場合やアドインの既定の設定値を入力するように求めるメッセージを表示する場合) にも、initialize イベント ハンドラーが使用できます。

 最小限の initialize イベントは、次の例のようになります。     

```js
Office.initialize = function () { };
```
独自の初期化ハンドラーやテストを含む追加の JavaScript フレームワークを使用している場合、そのようなフレームワークは Office.initialize イベントの内側に配置する必要があります。たとえば、[JQuery](https://jquery.com) の `$(document).ready()` 関数は次のように参照します。

```js
Office.initialize = function () {
    // Office is ready
    $(document).ready(function () {        
        // The document is ready
    });
  };
```
Office アドイン内のすべてのページで、初期化イベント **Office.initialize** にイベント ハンドラーを割り当てる必要があります。イベント ハンドラーを割り当てないと、アドインの起動時にエラーが発生することがあります。また、ユーザーが Excel Online、PowerPoint Online、Outlook Web App などの Office Online Web クライアントでアドインを使用しようとすると、アドインの実行が失敗します。初期化コードが必要ない場合は、上の最初の例のように、**Office.initialize** に割り当てる関数の本体を空にできます。

アドインの初期化時のイベントのシーケンスの詳細については、「[DOM とランタイム環境を読み込む](../develop/loading-the-dom-and-runtime-environment.md)」を参照してください。

#### <a name="initialization-reason"></a>初期化の理由
作業ウィンドウ アドインとコンテンツ アドインについては、Office.initialize に追加の _reason_ パラメーターを使用できます。このパラメーターは、アドインがどのように現在のドキュメントに追加されたかを判断するために使用できます。これは、最初にアドインが挿入されたときと、既にアドインがドキュメント内に存在しているときに別のロジックを提供するために使用できます。 

```js
Office.initialize = function (reason) {
    $(document).ready(function () {
      switch (reason) {
        case 'inserted': console.log('The add-in was just inserted.');
        case 'documentOpened': console.log('The add-in is already part of the document.');
    }
}
```
詳細については、「[Office.initialize イベント](http://dev.office.com/reference/add-ins/shared/office.initialize)」および「[InitializationReason 列挙型](http://dev.office.com/reference/add-ins/shared/initializationreason-enumeration)」を参照してください。 

## <a name="context-object"></a>Context オブジェクト

 **適用対象:** すべてのアドインの種類

アドインが初期化されると、多数のさまざまなオブジェクトとランタイム環境でやり取りできます。アドインのランタイム コンテキストは、[Context](http://dev.office.com/reference/add-ins/shared/office.context) オブジェクトによって API に反映されます。**Context** は中心的なオブジェクトで、これにより API の非常に重要なオブジェクト (たとえば、[Document](http://dev.office.com/reference/add-ins/shared/document) オブジェクトや [Mailbox](http://dev.office.com/reference/add-ins/outlook/Office.context.mailbox) オブジェクト) にアクセスできます。これらのオブジェクトにアクセスすると、次はドキュメントやメールボックスのコンテンツにアクセスできます。

たとえば、作業ウィンドウ アドインまたはコンテンツ アドインにおいて、[Context](http://dev.office.com/reference/add-ins/shared/office.context.document) オブジェクトの **document** プロパティを使用して、**Document** オブジェクトのプロパティおよびメソッドにアクセスし、Word 文書、Excel ワークシート、または Project スケジュールのコンテンツとやり取りできます。同様に、Outlook アドインにおいて、[Context](http://dev.office.com/reference/add-ins/outlook/Office.context.mailbox) オブジェクトの **mailbox** プロパティを使用して、**Mailbox** オブジェクトのプロパティおよびメソッドにアクセスし、メッセージ、会議出席依頼または予定のコンテンツとやり取りできます。

**Context** オブジェクトを使用すると、[contentLanguage](http://dev.office.com/reference/add-ins/shared/office.context.contentlanguage) プロパティと [displayLanguage](http://dev.office.com/reference/add-ins/shared/office.context.displaylanguage) プロパティにもアクセスできるようになり、ドキュメントやアイテム、またはホスト アプリケーションで使用するロケール (言語) を判断できます。また、[roamingSettings](http://dev.office.com/reference/add-ins/outlook/Office.context) プロパティを使用すると、[RoamingSettings](http://dev.office.com/reference/add-ins/outlook/RoamingSettings) オブジェクトのメンバーにアクセスできます。最後に、**Context** オブジェクトの [ui](http://dev.office.com/reference/add-ins/shared/officeui) プロパティを使用すると、アドインでポップアップ ダイアログを開始できます。


## <a name="document-object"></a>Document オブジェクト


 **適用対象:** コンテンツ アドインおよび作業ウィンドウ アドインの種類

Excel、PowerPoint、および Word のドキュメント データを操作するために、API には [Document](http://dev.office.com/reference/add-ins/shared/document) オブジェクトが用意されています。**Document** オブジェクトのメンバーを使用すると、次のようにデータにアクセスできます。


- テキスト、隣接するセル (マトリックス)、またはテーブルの形式のアクティブな選択範囲への読み取りと書き込み。
    
- 表形式のデータ (マトリックスまたはテーブル)。
    
- バインド (**Bindings** オブジェクトの "add" メソッドで作成)。
    
- カスタム XML パーツ (Word の場合のみ)。
    
- ドキュメント上のアドインごとに保持する設定またはアドインの状態。
    
また、**Document** オブジェクトを使用すると、Project ドキュメント内のデータを操作できます。API の Project 固有の機能については、[ProjectDocument](http://dev.office.com/reference/add-ins/shared/projectdocument.projectdocument) 抽象クラスのメンバー内に説明文があります。Project 用の作業ウィンドウ アドインの作成の詳細については、「[Project 用の作業ウィンドウ アドイン](../project/project-add-ins.md)」を参照してください。

これらのデータ アクセスの形式はすべて、抽象 **Document** オブジェクトのインスタンスから開始します。

作業ウィンドウ アドインまたはコンテンツ アドインが初期化されると、**Context** オブジェクトの [document](http://dev.office.com/reference/add-ins/shared/office.context.document) プロパティを使用して **Document** オブジェクトのインスタンスにアクセスできます。**Document** オブジェクトを使用すると、Word と Excel のドキュメントで共有される共通のデータ アクセス関数を定義でき、Word 文書の **CustomXmlParts** オブジェクトにもアクセスできます。

**Document** オブジェクトは、開発者がドキュメント コンテンツにアクセスするための 4 つの方法をサポートしています。


- 選択範囲ベースのアクセス
    
- バインドベースのアクセス
    
- カスタム XML パーツベースのアクセス (Word の場合のみ)
    
- ドキュメント全体へのアクセス (PowerPoint および Word のみ)
    
選択範囲ベースおよびバインドベースのデータ アクセス方法のしくみを理解するために、まず、データ アクセス API が、異なる Office アプリケーション間で一貫性のあるデータ アクセスを提供する方法について説明します。


### <a name="consistent-data-access-across-office-applications"></a>Office アプリケーション間での一貫性のあるデータ アクセス

 **適用対象:** コンテンツ アドインおよび作業ウィンドウ アドインの種類

異なる Office ドキュメント間でシームレスに動作する拡張機能を作成するために、JavaScript API for Office では、共通のデータ型と、異なるドキュメント コンテンツを 3 つの共通のデータ型に強制的に割り当てる機能を通じて、各 Office アプリケーションの特殊性を抽象化します。


#### <a name="common-data-types"></a>共通のデータ型

選択範囲ベースとバインドベースのどちらのデータ アクセスでも、ドキュメント コンテンツは、サポートされているすべての Office アプリケーション間で共通のデータ型を通じて公開されます。Office 2013 では、3 つの主要なデータ型がサポートされています。



|**データ型**|**説明**|**ホスト アプリケーションのサポート**|
|:-----|:-----|:-----|
|テキスト|選択範囲またはバインド内のデータの文字列表現を提供します。|Excel 2013、Project 2013、および PowerPoint 2013 は、プレーンテキストのみがサポートされます。Word 2013 では、3 つのテキスト形式 (プレーン テキスト、HTML、および Office Open XML (OOXML)) がサポートされます。Excel のセル内でテキストが選択されていると (セル内でテキストの一部のみが選択されている場合でも)、選択範囲ベースのメソッドは、セルのコンテンツ全体の読み取りおよび書き込みを行います。Word および PowerPoint でテキストが選択されていると、選択範囲ベースのメソッドは、選択されている文字の並びのみの読み取りおよび書き込みを行います。Project 2013 および PowerPoint 2013 は、選択範囲ベースのデータ アクセスのみをサポートします。|
|マトリックス|選択範囲またはバインドに含まれるデータを 2 次元の **Array** として提供します (JavaScript で配列の配列として実装されているものです)。たとえば、2 つの列にある 2 つ行の **string** 値は ` [['a', 'b'], ['c', 'd']]` になり、3 つの行を持つ 1 つの列は `[['a'], ['b'], ['c']]` になります。|マトリックス データ アクセスは Excel 2013 および Word 2013 でのみサポートされています。|
|テーブル|選択範囲またはバインド内のデータを [TableData](http://dev.office.com/reference/add-ins/shared/tabledata) オブジェクトとして提供します。**TableData** オブジェクトは、**headers** プロパティおよび **rows** プロパティを通じてデータを公開します。|テーブル データ アクセスは Excel 2013 および Word 2013 でのみサポートされています。|

#### <a name="data-type-coercion"></a>データ型の強制型変換

**Document** オブジェクトおよび [Binding](http://dev.office.com/reference/add-ins/shared/binding) オブジェクトのデータ アクセス メソッドでは、これらのメソッドの _coercionType_ パラメーターおよび対応する [CoercionType](http://dev.office.com/reference/add-ins/shared/coerciontype-enumeration) 列挙値を使用した目的のデータ型の指定をサポートしています。バインドの実際の形状にかかわらず、さまざまな Office アプリケーションでは、要求されるデータ型にデータを強制的に型変換することによって、共通のデータ型をサポートします。たとえば、Word の表または段落が選択されている場合、開発者はそれをプレーン テキスト、HTML、Office Open XML、または表として読み取ることを指定でき、API 実装によって必要な変換やデータ変換が行われます。


 >**ヒント**   **どのようなタイミングでデータ アクセスにマトリックスを使用し、どのような場合にテーブルの coercionType を使用するか。**行と列が追加されたときに表形式データが動的に増えるようにし、またテーブル ヘッダーを使用する必要がある場合は、テーブル データ型を使用します (**Document** または **Binding** オブジェクト データ アクセス メソッドの _coercionType_ パラメーターに `"table"` または **Office.CoercionType.Table** を指定)。データ構造体内での行と列の追加はテーブル データとマトリックス データの両方でサポートされていますが、行と列の追加はテーブル データでのみサポートされています。行と列を追加する予定がなく、データにヘッダー機能が必要ない場合は、マトリックス データ型を使用します (データ アクセス メソッドの _coercionType_ パラメーターに `"matrix"` または **Office.CoercionType.Matrix** を指定)。このデータ型では、データとのやり取りについて、より単純なモデルを採用しています。

指定された型にデータを強制的に型変換できない場合は、コールバック内の [AsyncResult.status](http://dev.office.com/reference/add-ins/shared/asyncresult.error) プロパティが `"failed"` を返すため、[AsyncResult.error](http://dev.office.com/reference/add-ins/shared/asyncresult.context) プロパティを使用して [Error](http://dev.office.com/reference/add-ins/shared/error) オブジェクトにアクセスし、メソッド呼び出しが失敗した理由を確認できます。


## <a name="working-with-selections-using-the-document-object"></a>Document オブジェクトによる選択範囲の操作


**Document** オブジェクトは、ユーザーの現在の選択を「取得および設定」の方法で読み取りおよび書き込みできるメソッドを公開します。そのために、**Document** オブジェクトは **getSelectedDataAsync** メソッドと **setSelectedDataAsync** メソッドを提供します。

選択範囲に関する操作の実行方法を示すコード例については、「[ドキュメントまたはスプレッドシート内のアクティブな選択範囲へのデータの読み取りおよび書き込み](../develop/read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md)」を参照してください。


## <a name="working-with-bindings-using-the-bindings-and-binding-objects"></a>Bindings オブジェクトおよび Binding オブジェクトによるバインドの操作


バインドベースのデータ アクセスを使用すると、コンテンツ アドインおよび作業ウィンドウ アドインで、バインドに関連付けられた識別子を介して、ドキュメントまたはスプレッドシートの特定の領域に一貫性のあるアクセスが可能になります。アドインは、最初に、ドキュメントの部分と一意の ID を関連付けるメソッドのいずれか ([addFromPromptAsync](http://dev.office.com/reference/add-ins/shared/bindings.addfrompromptasync)、[addFromSelectionAsync](http://dev.office.com/reference/add-ins/shared/bindings.addfromselectionasync)、または [addFromNamedItemAsync](http://dev.office.com/reference/add-ins/shared/bindings.addfromnameditemasync)) を呼び出すことによって、バインドを確立する必要があります。バインドが確立されると、アドインは提供された ID を使用して、ドキュメントまたはスプレッドシート内の関連付けられた領域に含まれるデータにアクセスできます。バインドを作成すると、アドインには次のようなメリットがあります。


- 表、範囲、またはテキスト (隣接する一連の文字) など、サポートされている Office アプリケーション全体に共通のデータ構造へのアクセスを許可します。
    
- ユーザーによる選択を必要とせずに、読み取り/書き込み操作ができます。
    
- アドインとドキュメント内のデータの間にリレーションシップが確立されます。バインドはドキュメント内に保持され、後でアクセスできます。
    
また、バインドを確立すると、ドキュメントまたはスプレッドシートの特定の領域を範囲とする、データおよび選択範囲の変更イベントをサブスクライブできます。つまり、ドキュメントまたはスプレッドシート全体の全般的な変更ではなく、バインドされた領域内で発生する変更のみがアドインに通知されます。

[Bindings](http://dev.office.com/reference/add-ins/shared/bindings.bindings) オブジェクトが公開している [getAllAsync](http://dev.office.com/reference/add-ins/shared/bindings.getallasync) メソッドを使用すると、ドキュメントまたはスプレッドシートで確立されている一連のすべてのバインドにアクセスできます。個々のバインドに ID でアクセスするには、[Bindings.getBindingByIdAsync](http://dev.office.com/reference/add-ins/shared/bindings.getbyidasync) メソッドまたは [Office.select](http://dev.office.com/reference/add-ins/shared/office.select) メソッドを使用します。**Bindings** オブジェクトのいずれかのメソッド ([addFromSelectionAsync](http://dev.office.com/reference/add-ins/shared/bindings.addfromselectionasync)、[addFromPromptAsync](http://dev.office.com/reference/add-ins/shared/bindings.addfrompromptasync)、[addFromNamedItemAsync](http://dev.office.com/reference/add-ins/shared/bindings.addfromnameditemasync)、または [releaseByIdAsync](http://dev.office.com/reference/add-ins/shared/bindings.releasebyidasync)) を使用すると、新しいバインドを確立したり既存のバインドを削除したりできます。

_addFromSelectionAsync_ メソッド、**addFromPromptAsync** メソッド、または **addFromNamedItemAsync** メソッドでバインドを作成する場合、**bindingType** パラメーターで指定するバインドには 3 つの種類あります。



|**バインドの種類**|**説明**|**ホスト アプリケーションのサポート**|
|:-----|:-----|:-----|
|テキスト バインド|テキストとして表現できるドキュメントの領域にバインドします。|Word では、連続する選択範囲の大部分が有効ですが、Excel では、単一セルの範囲のみがテキスト バインドの対象です。Excel では、プレーン テキストのみがサポートされます。Word では、3 つの形式 (プレーン テキスト、HTML、および Open XML for Office) がサポートされます。|
|マトリックス バインド|ヘッダーがない表形式のデータが含まれるドキュメントの固定領域にバインドします。マトリックス バインド内のデータは、2 次元の  **Array** として書き込みまたは読み取りが行われます。JavaScript では、これは、配列の配列として実装されています。たとえば、2 列の **string** 値が 2 行ある場合は ` [['a', 'b'], ['c', 'd']]` のように書き込みまたは読み取りが行われ、1 列が 3 行ある場合は `[['a'], ['b'], ['c']]` のように書き込みまたは読み取りが行われます。|Excel では、セルの連続する選択範囲を使用してマトリックス バインドを確立できます。Word では、表のみがマトリックス バインドをサポートします。|
|テーブル バインド|ヘッダーがある表が含まれるドキュメントの領域にバインドします。テーブル バインド内のデータは、[TableData](http://dev.office.com/reference/add-ins/shared/tabledata) オブジェクトとして書き込みまたは読み取りが行われます。**TableData** オブジェクトは **headers** および **rows** プロパティを通じてデータを公開します。|Excel または Word の表はすべて、テーブル バインドの基礎にできます。テーブル バインドを確立すると、ユーザーが表に追加する新しい各行または各列が、自動的にバインドに含まれます。 |
**Bindings** オブジェクトの 3 つの "add" メソッドのいずれかを使用してバインドを作成すると、[MatrixBinding](http://dev.office.com/reference/add-ins/shared/binding.matrixbinding)、[TableBinding](http://dev.office.com/reference/add-ins/shared/binding.tablebinding)、または [TextBinding](http://dev.office.com/reference/add-ins/shared/binding.textbinding) のうち対応するオブジェクトのメソッドを使用して、バインドのデータおよびプロパティを操作できます。この 3 つのオブジェクトはすべて、[Binding](http://dev.office.com/reference/add-ins/shared/binding.getdataasync) オブジェクトの [getDataAsync](http://dev.office.com/reference/add-ins/shared/binding.setdataasync) メソッドおよび **setDataAsync** メソッドを継承しているので、バインドされたデータを操作できます。

バインドに関する操作の実行方法を示すコード例については、「[ドキュメントまたはスプレッドシート内の領域へのバインド](../develop/bind-to-regions-in-a-document-or-spreadsheet.md)」を参照してください。


## <a name="working-with-custom-xml-parts-using-the-customxmlparts-and-customxmlpart-objects"></a>CustomXmlParts オブジェクトおよび CustomXmlPart オブジェクトによるカスタム XML パーツの操作


 **適用対象:** Word の作業ウィンドウ アドイン

API の [CustomXmlParts](http://dev.office.com/reference/add-ins/shared/customxmlparts.customxmlparts) オブジェクトと [CustomXmlPart](http://dev.office.com/reference/add-ins/shared/customxmlpart.customxmlpart) オブジェクトを使用すると、Word 文書内のカスタム XML パーツにアクセスできます。これにより、文書のコンテンツに対する XML 主導の操作が可能になります。**CustomXmlParts** オブジェクトおよび **CustomXmlPart** オブジェクトとの連携のデモについては、「[Word-Add-in-Work-with-custom-XML-parts](https://github.com/OfficeDev/Word-Add-in-Work-with-custom-XML-parts)」のコード例を参照してください。


## <a name="working-with-the-entire-document-using-the-getfileasync-method"></a>getFileAsync メソッドを使用したドキュメント全体の操作


 **適用対象:** Word および PowerPoint の作業ウィンドウ アドイン

[Document.getFileAsync](http://dev.office.com/reference/add-ins/shared/document.getfileasync) メソッド、および [File](http://dev.office.com/reference/add-ins/shared/file) オブジェクトと [Slice](http://dev.office.com/reference/add-ins/shared/slice) オブジェクトのメンバーは、一度に最大で 4 MB ずつのスライス (チャンク) に分割して Word および PowerPoint ドキュメント ファイル全体を取得する機能を提供します。詳細については、「[アドインのドキュメントですべてのファイルを取得する方法](../develop/get-the-whole-document-from-an-add-in-for-powerpoint-or-word.md)」を参照してください。


## <a name="mailbox-object"></a>Mailbox オブジェクト


 **適用対象:** Outlook アドイン

Outlook アドインでは、主に [Mailbox](http://dev.office.com/reference/add-ins/outlook/Office.context.mailbox) オブジェクトにより公開されている API のサブセットを使用します。Outlook アドイン専用のオブジェクトおよびメンバー (たとえば、[Item](http://dev.office.com/reference/add-ins/outlook/Office.context.mailbox.item) オブジェクトなど) にアクセスするには、次のコード行に示すように、[Context](http://dev.office.com/reference/add-ins/outlook/Office.context.mailbox) オブジェクトの **mailbox** プロパティを使用して、**Mailbox** オブジェクトにアクセスします。




```js
// Access the Item object.
var item = Office.context.mailbox.item;

```

さらに、Outlook アドインでは次のオブジェクトを使用できます。


-  **Office** オブジェクト: 初期化に使用します。
    
-  **Context** オブジェクト: コンテンツおよび表示言語のプロパティへのアクセスに使用します。
    
-  **RoamingSettings** オブジェクト: アドインがインストールされているユーザーのメールボックスに Outlook アドイン固有のカスタム設定を保存する際に使用します。
    
Outlook アドインで JavaScript を使用する方法の詳細については、「[Outlook アドイン](../outlook/outlook-add-ins.md)」および「[Outlook アドインのアーキテクチャと機能の概要](../outlook/overview.md)」を参照してください。


## <a name="api-support-matrix"></a>API サポート マトリックス


この表は、アドインの種類 (コンテンツ、作業ウィンドウ、および Outlook) 全体でサポートされている API と機能、および [1.1 アドイン マニフェスト スキーマと機能 (JavaScript API for Office v1.1 でサポート)](http://msdn.microsoft.com/library/cff9fbdf-a530-4f6e-91ca-81bcacd90dcd%28Office.15%29.aspx) を使用して [アドインがサポートする Office のホスト アプリケーション](../develop/update-your-javascript-api-for-office-and-manifest-schema-version.md) を指定する際にこれらの API と機能をホストする Office アプリケーションについてまとめたものです。


|||||||||
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
||**ホスト名**|データベース|ブック|メールボックス|プレゼンテーション|ドキュメント|Project|
||**サポートされる****ホスト アプリケーション**|Access Web アプリ|Excel、Excel Online|Outlook、Outlook Web App、デバイス用 OWA|PowerPoint、PowerPoint Online|Word|プロジェクト|
|**サポートされるアドインの種類**|コンテンツ|Y|Y||Y|||
||作業ウィンドウ||Y||Y|Y|Y|
||Outlook|||Y||||
|**サポートされている API 機能**|テキストの読み取り/書き込み||Y||Y|Y|Y (読み取り専用)|
||マトリックスの読み取り/書き込み||Y|||Y||
||テーブルの読み取り/書き込み||Y|||Y||
||HTML の読み取り/書き込み|||||Y||
||Office Open XML の読み取り/書き込み|||||Y||
||タスク、リソース、ビュー、フィールド プロパティの読み取り||||||Y|
||選択変更イベント||Y|||Y||
||ドキュメント全体の取得||||Y|Y||
||バインドとイベント バインド|Y (完全なテーブル バインドと部分的なテーブル バインドのみ)|Y|||Y||
||カスタム XML パーツの読み取り/書き込み|||||Y||
||アドイン状態データの保持 (設定)|Y (ホスト アドインごと)|Y (ドキュメントごと)|Y (メールボックスごと)|Y (ドキュメントごと)|Y (ドキュメントごと)||
||設定変更イベント|Y|Y||Y|Y||
||アクティブ ビュー モードおよびビュー変更イベントの取得||||Y|||
||ドキュメント内の場所に移動||Y||Y|Y||
||ルールと RegEx を使用した文脈からのアクティブ化|||Y||||
||アイテム プロパティの読み取り|||Y||||
||ユーザー プロファイルの読み取り|||Y||||
||添付ファイルの取得|||Y||||
||ユーザー ID トークンの取得|||Y||||
||Exchange Web サービスの呼出|||Y||||
