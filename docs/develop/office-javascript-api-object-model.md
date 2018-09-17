---
title: Office JavaScript API オブジェクト モデル
description: ''
ms.date: 07/27/2018
ms.openlocfilehash: 0842d9deafd8df411f3074dcddca04ebe0f9ed02
ms.sourcegitcommit: 30435939ab8b8504c3dbfc62fd29ec6b0f1a7d22
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/12/2018
ms.locfileid: "23945675"
---
# <a name="office-javascript-api-object-model"></a>Office JavaScript API オブジェクト モデル
Office JavaScript アドインは、ホストの基本機能にアクセスできるようにします。 このアクセスのほとんどは、いくつかの重要なオブジェクトを経由します。 [Context](#context-object) オブジェクトは、初期化後にランタイム環境にアクセスします。 [Document](#document-object) オブジェクトを使用すると、ユーザーはExcel、PowerPoint、または Word ドキュメントを制御できます。 [Mailbox](#mailbox-object) オブジェクトは、メッセージとユーザー プロファイルへの Outlook アドイン アクセスを提供します。 これらの高水準オブジェクト間の関係を理解することは、JavaScript アドインの基礎となります。

## <a name="context-object"></a>Context オブジェクト

**適用対象:** すべてのアドインの種類

アドインが [初期化される](https://docs.microsoft.com/office/dev/add-ins/develop/understanding-the-javascript-api-for-office#initializing-your-add-in)と、多数のさまざまなオブジェクトとランタイム環境でやり取りできます。 アドインのランタイム コンテキストは、 [Context](https://docs.microsoft.com/javascript/api/office/office.context?view=office-js) オブジェクトによって API に反映されます。 **Context** は API の最も重要なオブジェクトへのアクセスを提供する主要オブジェクトです。[Document](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js) と [Mailbox](https://docs.microsoft.com/javascript/api/outlook/Office.mailbox?view=office-js) オブジェクトは、ドキュメントとメールボックスのコンテンツへのアクセスを提供します。

たとえば、作業ウィンドウ アドインまたはコンテンツ アドインにおいて、[Context](https://docs.microsoft.com/javascript/api/office/office.context?view=office-js#document) オブジェクトの **document** プロパティを使用して、**Document** オブジェクトのプロパティおよびメソッドにアクセスし、Word 文書、Excel ワークシート、または Project スケジュールのコンテンツとやり取りできます。同様に、Outlook アドインにおいて、[Context](https://docs.microsoft.com/javascript/api/outlook/Office.mailbox?view=office-js) オブジェクトの **mailbox** プロパティを使用して、**Mailbox** オブジェクトのプロパティおよびメソッドにアクセスし、メッセージ、会議出席依頼または予定のコンテンツとやり取りできます。

**Context** オブジェクトを使用すると、[contentLanguage](https://docs.microsoft.com/javascript/api/office/office.context?view=office-js#contentlanguage) プロパティと [displayLanguage](https://docs.microsoft.com/javascript/api/office/office.context?view=office-js#displaylanguage) プロパティにもアクセスが可能になり、ドキュメントやアイテム、またはホスト アプリケーションで使用するロケール (言語) を判断できます。 [roamingSettings](https://docs.microsoft.com/javascript/api/office/office.context?view=office-js#roamingsettings) プロパティを使用すると、 [RoamingSettings](https://docs.microsoft.com/javascript/api/office/office.context?view=office-js#roamingsettings) オブジェクトは、個々のユーザーのメールボックスのアドインに固有の設定を格納します。 最後に、**Context** オブジェクトの [ui](https://docs.microsoft.com/javascript/api/office/office.ui?view=office-js) プロパティを使用すると、アドインでポップアップ ダイアログを開始できます。


## <a name="document-object"></a>Document オブジェクト

**適用対象:** コンテンツ アドインおよび作業ウィンドウ アドインの種類

Excel、PowerPoint、および Word のドキュメント データを操作するために、API には [Document](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js) オブジェクトが用意されています。**Document** オブジェクトのメンバーを使用すると、次のようにデータにアクセスできます。

- テキスト、隣接するセル (マトリックス)、またはテーブルの形式のアクティブな選択範囲への読み取りと書き込み。
    
- 表形式のデータ (マトリックスまたはテーブル)。
    
- バインド (**Bindings** オブジェクトの "add" メソッドで作成)。
    
- カスタム XML パーツ (Word の場合のみ)。
    
- ドキュメント上のアドインごとに保持する設定またはアドインの状態。
    
また、**Document** オブジェクトを使用すると、Project ドキュメント内のデータを操作できます。API の Project 固有の機能については、[ProjectDocument](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js) 抽象クラスのメンバー内に説明文があります。Project 用の作業ウィンドウ アドインの作成の詳細については、「[Project 用の作業ウィンドウ アドイン](../project/project-add-ins.md)」を参照してください。

これらのデータ アクセスの形式はすべて、抽象 **Document** オブジェクトのインスタンスから開始します。

作業ウィンドウ アドインまたはコンテンツ アドインが初期化されると、**Context** オブジェクトの [document](https://docs.microsoft.com/javascript/api/office/office.context?view=office-js#document) プロパティを使用して **Document** オブジェクトのインスタンスにアクセスできます。**Document** オブジェクトを使用すると、Word と Excel のドキュメントで共有される共通のデータ アクセス関数を定義でき、Word 文書の **CustomXmlParts** オブジェクトにもアクセスできます。

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
|テーブル|選択範囲またはバインド内のデータを [TableData](https://docs.microsoft.com/javascript/api/office/office.tabledata?view=office-js) オブジェクトとして提供します。**TableData** オブジェクトは、**headers** プロパティおよび **rows** プロパティを通じてデータを公開します。|テーブル データ アクセスは Excel 2013 および Word 2013 でのみサポートされています。|

#### <a name="data-type-coercion"></a>データ型の強制型変換

**Document** オブジェクトおよび [Binding](https://docs.microsoft.com/javascript/api/office/office.binding?view=office-js) オブジェクトのデータ アクセス メソッドでは、これらのメソッドの _coercionType_ パラメーターおよび対応する [CoercionType](https://docs.microsoft.com/javascript/api/office/office.coerciontype?view=office-js) 列挙値を使用した目的のデータ型の指定をサポートしています。バインドの実際の形状にかかわらず、さまざまな Office アプリケーションでは、要求されるデータ型にデータを強制的に型変換することによって、共通のデータ型をサポートします。たとえば、Word の表または段落が選択されている場合、開発者はそれをプレーン テキスト、HTML、Office Open XML、または表として読み取ることを指定でき、API 実装によって必要な変換やデータ変換が行われます。


> [!TIP]
> **どのようなタイミングでデータ アクセスにマトリックスを使用し、どのような場合にテーブルの coercionType を使用するか。** 行と列が追加されたときに表形式データが動的に増えるようにし、またテーブル ヘッダーを使用する必要がある場合は、テーブル データ型を使用します (**Document** または **Binding** オブジェクト データ アクセス メソッドの _coercionType_ パラメーターに `"table"` または **Office.CoercionType.Table** を指定)。データ構造体内での行と列の追加はテーブル データとマトリックス データの両方でサポートされていますが、行と列の追加はテーブル データでのみサポートされています。行と列を追加する予定がなく、データにヘッダー機能が必要ない場合は、マトリックス データ型を使用します (データ アクセス メソッドの _coercionType_ パラメーターに `"matrix"` または **Office.CoercionType.Matrix** を指定)。このデータ型では、データとのやり取りについて、より単純なモデルを採用しています。

指定された型にデータを強制的に型変換できない場合は、コールバック内の [AsyncResult.status](https://docs.microsoft.com/javascript/api/office/office.asyncresult?view=office-js.error) プロパティが `"failed"` を返すため、[AsyncResult.error](https://docs.microsoft.com/javascript/api/office/office.asyncresult?view=office-js.context) プロパティを使用して [Error](https://docs.microsoft.com/javascript/api/office/office.error?view=office-js) オブジェクトにアクセスし、メソッド呼び出しが失敗した理由を確認できます。


## <a name="working-with-selections-using-the-document-object"></a>Document オブジェクトによる選択範囲の操作


**Document** オブジェクトは、ユーザーの現在の選択を「取得および設定」の方法で読み取りおよび書き込みできるメソッドを公開します。そのために、**Document** オブジェクトは **getSelectedDataAsync** メソッドと **setSelectedDataAsync** メソッドを提供します。

選択範囲に関する操作の実行方法を示すコード例については、「[ドキュメントまたはスプレッドシート内のアクティブな選択範囲へのデータの読み取りおよび書き込み](read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md)」を参照してください。


## <a name="working-with-bindings-using-the-bindings-and-binding-objects"></a>Bindings オブジェクトおよび Binding オブジェクトによるバインドの操作


バインドベースのデータ アクセスを使用すると、コンテンツ アドインおよび作業ウィンドウ アドインで、バインドに関連付けられた識別子を介して、ドキュメントまたはスプレッドシートの特定の領域に一貫性のあるアクセスが可能になります。アドインは、最初に、ドキュメントの部分と一意の ID を関連付けるメソッドのいずれか ([addFromPromptAsync](https://docs.microsoft.com/javascript/api/office/office.bindings?view=office-js#addfrompromptasync-bindingtype--options--callback-)、[addFromSelectionAsync](https://docs.microsoft.com/javascript/api/office/office.bindings?view=office-js#addfromselectionasync-bindingtype--options--callback-)、または [addFromNamedItemAsync](https://docs.microsoft.com/javascript/api/office/office.bindings?view=office-js#addfromnameditemasync-itemname--bindingtype--options--callback-)) を呼び出すことによって、バインドを確立する必要があります。バインドが確立されると、アドインは提供された ID を使用して、ドキュメントまたはスプレッドシート内の関連付けられた領域に含まれるデータにアクセスできます。バインドを作成すると、アドインには次のようなメリットがあります。


- 表、範囲、またはテキスト (隣接する一連の文字) など、サポートされている Office アプリケーション全体に共通のデータ構造へのアクセスを許可します。
    
- ユーザーによる選択を必要とせずに、読み取り/書き込み操作ができます。
    
- アドインとドキュメント内のデータの間にリレーションシップが確立されます。バインドはドキュメント内に保持され、後でアクセスできます。
    
また、バインドを確立すると、ドキュメントまたはスプレッドシートの特定の領域を範囲とする、データおよび選択範囲の変更イベントをサブスクライブできます。つまり、ドキュメントまたはスプレッドシート全体の全般的な変更ではなく、バインドされた領域内で発生する変更のみがアドインに通知されます。

[Bindings](https://docs.microsoft.com/javascript/api/office/office.bindings?view=office-js) オブジェクトが公開している [getAllAsync](https://docs.microsoft.com/javascript/api/office/office.bindings?view=office-js#getallasync-options--callback-) メソッドを使用すると、ドキュメントまたはスプレッドシートで確立されている一連のすべてのバインドにアクセスできます。個々のバインドに ID でアクセスするには、[Bindings.getBindingByIdAsync](https://docs.microsoft.com/javascript/api/office/office.bindings?view=office-js#getbyidasync-id--options--callback-) メソッドまたは [Office.select](https://docs.microsoft.com/javascript/api/office?view=office-js) メソッドを使用します。**Bindings** オブジェクトのいずれかのメソッド ([addFromSelectionAsync](https://docs.microsoft.com/javascript/api/office/office.bindings?view=office-js#addfromselectionasync-bindingtype--options--callback-)、[addFromPromptAsync](https://docs.microsoft.com/javascript/api/office/office.bindings?view=office-js#addfrompromptasync-bindingtype--options--callback-)、[addFromNamedItemAsync](https://docs.microsoft.com/javascript/api/office/office.bindings?view=office-js#addfromnameditemasync-itemname--bindingtype--options--callback-)、または [releaseByIdAsync](https://docs.microsoft.com/javascript/api/office/office.bindings?view=office-js#releasebyidasync-id--options--callback-)) を使用すると、新しいバインドを確立したり既存のバインドを削除したりできます。

_addFromSelectionAsync_ メソッド、**addFromPromptAsync** メソッド、または **addFromNamedItemAsync** メソッドでバインドを作成する場合、**bindingType** パラメーターで指定するバインドには 3 つの種類あります。



|**バインドの種類**|**説明**|**ホスト アプリケーションのサポート**|
|:-----|:-----|:-----|
|テキスト バインド|テキストとして表現できるドキュメントの領域にバインドします。|Word では、連続する選択範囲の大部分が有効ですが、Excel では、単一セルの範囲のみがテキスト バインドの対象です。Excel では、プレーン テキストのみがサポートされます。Word では、3 つの形式 (プレーン テキスト、HTML、および Open XML for Office) がサポートされます。|
|マトリックス バインド|ヘッダーがない表形式のデータが含まれるドキュメントの固定領域にバインドします。マトリックス バインド内のデータは、2 次元の **Array** として書き込みまたは読み取りが行われます。JavaScript では、これは、配列の配列として実装されています。たとえば、2 列の **string** 値が 2 行ある場合は ` [['a', 'b'], ['c', 'd']]` のように書き込みまたは読み取りが行われ、1 列が 3 行ある場合は `[['a'], ['b'], ['c']]` のように書き込みまたは読み取りが行われます。|Excel では、セルの連続する選択範囲を使用してマトリックス バインドを確立できます。Word では、表のみがマトリックス バインドをサポートします。|
|テーブル バインド|ヘッダーがある表が含まれるドキュメントの領域にバインドします。テーブル バインド内のデータは、[TableData](https://docs.microsoft.com/javascript/api/office/office.tabledata?view=office-js) オブジェクトとして書き込みまたは読み取りが行われます。**TableData** オブジェクトは **headers** および **rows** プロパティを通じてデータを公開します。|Excel または Word の表はすべて、テーブル バインドの基礎にできます。テーブル バインドを確立すると、ユーザーが表に追加する新しい各行または各列が、自動的にバインドに含まれます。 |

<br/>

**Bindings** オブジェクトの 3 つの "add" メソッドのいずれかを使用してバインドを作成すると、[MatrixBinding](https://docs.microsoft.com/javascript/api/office/office.matrixbinding?view=office-js)、[TableBinding](https://docs.microsoft.com/javascript/api/office/office.tablebinding?view=office-js)、または [TextBinding](https://docs.microsoft.com/javascript/api/office/office.textbinding?view=office-js) のうち対応するオブジェクトのメソッドを使用して、バインドのデータおよびプロパティを操作できます。この 3 つのオブジェクトはすべて、[Binding](https://docs.microsoft.com/javascript/api/office/office.binding?view=office-js#getdataasync-options--callback-) オブジェクトの [getDataAsync](https://docs.microsoft.com/javascript/api/office/office.binding?view=office-js#setdataasync-data--options--callback-) メソッドおよび **setDataAsync** メソッドを継承しているので、バインドされたデータを操作できます。

バインドに関する操作の実行方法を示すコード例については、「[ドキュメントまたはスプレッドシート内の領域へのバインド](bind-to-regions-in-a-document-or-spreadsheet.md)」を参照してください。


## <a name="working-with-custom-xml-parts-using-the-customxmlparts-and-customxmlpart-objects"></a>CustomXmlParts オブジェクトおよび CustomXmlPart オブジェクトによるカスタム XML パーツの操作


 **適用対象:** Word の作業ウィンドウ アドイン

API の [CustomXmlParts](https://docs.microsoft.com/javascript/api/office/office.customxmlparts?view=office-js) オブジェクトと [CustomXmlPart](https://docs.microsoft.com/javascript/api/office/office.customxmlpart?view=office-js) オブジェクトを使用すると、Word 文書内のカスタム XML パーツにアクセスできます。これにより、文書のコンテンツに対する XML 主導の操作が可能になります。**CustomXmlParts** オブジェクトおよび **CustomXmlPart** オブジェクトとの連携のデモについては、「[Word-Add-in-Work-with-custom-XML-parts](https://github.com/OfficeDev/Word-Add-in-Work-with-custom-XML-parts)」のコード例を参照してください。


## <a name="working-with-the-entire-document-using-the-getfileasync-method"></a>getFileAsync メソッドを使用したドキュメント全体の操作


 **適用対象:** Word および PowerPoint の作業ウィンドウ アドイン

[Document.getFileAsync](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#getfileasync-filetype--options--callback-) メソッド、および [File](https://docs.microsoft.com/javascript/api/office/office.file?view=office-js) オブジェクトと [Slice](https://docs.microsoft.com/javascript/api/office/office.slice?view=office-js) オブジェクトのメンバーは、一度に最大で 4 MB ずつのスライス (チャンク) に分割して Word および PowerPoint ドキュメント ファイル全体を取得する機能を提供します。詳細については、「[PowerPoint または Word 用アドインからドキュメント全体を取得する](../word/get-the-whole-document-from-an-add-in-for-word.md)」を参照してください。


## <a name="mailbox-object"></a>Mailbox オブジェクト

**適用対象:** Outlook アドイン

Outlook アドインでは、主に [Mailbox](https://docs.microsoft.com/javascript/api/outlook/Office.mailbox?view=office-js) オブジェクトにより公開されている API のサブセットを使用します。Outlook アドイン専用のオブジェクトおよびメンバー (たとえば、[Item](https://docs.microsoft.com/javascript/api/outlook/Office.mailbox?view=office-js) オブジェクトなど) にアクセスするには、次のコード行に示すように、[Context](https://docs.microsoft.com/javascript/api/outlook/Office.mailbox?view=office-js) オブジェクトの **mailbox** プロパティを使用して、**Mailbox** オブジェクトにアクセスします。




```js
// Access the Item object.
var item = Office.context.mailbox.item;

```

さらに、Outlook アドインでは次のオブジェクトを使用できます。


-  **Office** オブジェクト: 初期化に使用します。
    
-  **Context** オブジェクト: コンテンツおよび表示言語のプロパティへのアクセスに使用します。
    
-  **RoamingSettings** オブジェクト: アドインがインストールされているユーザーのメールボックスに Outlook アドイン固有のカスタム設定を保存する際に使用します。
    
Outlook アドインでの JavaScript の使用については、「[Outlook アドイン](https://docs.microsoft.com/outlook/add-ins/)」を参照してください。