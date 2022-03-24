---
title: 共通 JavaScript API オブジェクト モデル
description: JavaScript 共通 API Officeモデルについて説明します。
ms.date: 07/08/2021
ms.localizationpriority: medium
ms.openlocfilehash: 381d3089b47fe04f403459ecae249bf68f7ca872
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/23/2022
ms.locfileid: "63743773"
---
# <a name="common-javascript-api-object-model"></a>共通 JavaScript API オブジェクト モデル

[!include[information about the common API](../includes/alert-common-api-info.md)]

Office JavaScript API を使用すると、クライアント アプリケーションOffice基になる機能にアクセスできます。 このアクセスの大部分はいくつかの重要なオブジェクトを通過します。 [Context](#context-object) オブジェクトによって、初期化した後、ランタイム環境にアクセスできるようになります。 [Document](#document-object) オブジェクトによって、Excel、PowerPoint、Word ドキュメントを操作する許可が与えられます。 [Mailbox オブジェクト](#mailbox-object)は、Outlook、予定、およびユーザー プロファイルへのアクセス権を提供します。 これらの高レベル オブジェクト間の関係を理解すると、アドインのOffice基礎になります。

## <a name="context-object"></a>Context オブジェクト

**適用対象:** すべてのアドインの種類

アドインが[初期化](initialize-add-in.md)されると、ランタイム環境でやり取りできるさまざまなオブジェクトが多数あります。 アドインのランタイム コンテキストは [Context](/javascript/api/office/office.context) オブジェクトによって API で反映されます。 **Context** は、[Document](/javascript/api/office/office.document) オブジェクトや [Mailbox](/javascript/api/outlook/office.mailbox) オブジェクトなど、API の最重要オブジェクトにアクセスできるメイン オブジェクトです。最重要オブジェクトはさらにドキュメントやメールボックスのコンテンツにアクセスできます。

たとえば、作業ウィンドウ アドインまたはコンテンツ アドインにおいて、[Context](/javascript/api/office/office.context#office-office-context-document-member) オブジェクトの **document** プロパティを使用して、**Document** オブジェクトのプロパティおよびメソッドにアクセスし、Word 文書、Excel ワークシート、または Project スケジュールのコンテンツとやり取りできます。同様に、Outlook アドインにおいて、[Context](/javascript/api/office/office.context#office-office-context-mailbox-member) オブジェクトの **mailbox** プロパティを使用して、**Mailbox** オブジェクトのプロパティおよびメソッドにアクセスし、メッセージ、会議出席依頼または予定のコンテンツとやり取りできます。

**Context** オブジェクトでは、[contentLanguage プロパティと displayLanguage](/javascript/api/office/office.context#office-office-context-contentlanguage-member) プロパティにアクセスして、ドキュメントまたはアイテム、または Office アプリケーションで使用されるロケール (言語) を特定できます。[](/javascript/api/office/office.context#office-office-context-displaylanguage-member) [roamingSettings](/javascript/api/office/office.context#office-office-context-roamingsettings-member) プロパティによって、[RoamingSettings](/javascript/api/office/office.context#office-office-context-roamingsettings-member) オブジェクトのメンバーにアクセスできます。このオブジェクトによって、個々のユーザーのメールボックスに対してアドインに固有の設定が保存されます。 最後に、**Context** オブジェクトの [ui](/javascript/api/office/office.context#office-office-context-ui-member) プロパティを使用すると、アドインでポップアップ ダイアログを開始できます。

## <a name="document-object"></a>Document オブジェクト

**適用対象:** コンテンツ アドインおよび作業ウィンドウ アドインの種類

Excel、PowerPoint、および Word のドキュメント データを操作するために、API には [Document](/javascript/api/office/office.document) オブジェクトが用意されています。 オブジェクト メンバーを使用 `Document` すると、次の方法でデータにアクセスできます。

- テキスト、隣接するセル (マトリックス)、またはテーブルの形式のアクティブな選択範囲への読み取りと書き込み。

- 表形式のデータ (マトリックスまたはテーブル)。

- バインド (オブジェクトの "add" メソッドを使用して作成 `Bindings` されます)。

- カスタム XML パーツ (Word の場合のみ)。

- ドキュメント上のアドインごとに保持する設定またはアドインの状態。

オブジェクトを使用して、`Document`ドキュメント内のデータをProjectできます。 API の Project 固有の機能については、[ProjectDocument](/javascript/api/office/office.document) 抽象クラスのメンバー内に説明文があります。 Project 用の作業ウィンドウ アドインの作成の詳細については、「[Project 用の作業ウィンドウ アドイン](../project/project-add-ins.md)」を参照してください。

これらのデータ アクセスのすべての形式は、抽象オブジェクトのインスタンスから始 `Document` まる。

オブジェクトの document プロパティを `Document` 使用して、作業ウィンドウアドインまたはコンテンツ アドインを初期化するときに、 [オブジェクトのインスタンス](/javascript/api/office/office.context#office-office-context-document-member) にアクセス `Context` できます。 この`Document`オブジェクトは、Word およびドキュメント間で共有される一般的なデータ アクセス`CustomXmlParts`Excel定義し、Word ドキュメントのオブジェクトへのアクセスも提供します。

このオブジェクト `Document` は、開発者がドキュメントコンテンツにアクセスするための 4 つの方法をサポートしています。

- 選択範囲ベースのアクセス

- バインドベースのアクセス

- カスタム XML パーツベースのアクセス (Word の場合のみ)

- ドキュメント全体へのアクセス (PowerPoint および Word のみ)

選択範囲ベースおよびバインドベースのデータ アクセス方法のしくみを理解するために、まず、データ アクセス API が、異なる Office アプリケーション間で一貫性のあるデータ アクセスを提供する方法について説明します。

### <a name="consistent-data-access-across-office-applications"></a>Office アプリケーション間での一貫性のあるデータ アクセス

 **適用対象:** コンテンツ アドインおよび作業ウィンドウ アドインの種類

Office JavaScript API は、異なる Office ドキュメント間でシームレスに機能する拡張機能を作成するために、共通のデータ型を使用して各 Office アプリケーションの特定性を抽象化し、異なるドキュメントコンテンツを 3 つの一般的なデータ型に変換できます。

#### <a name="common-data-types"></a>共通のデータ型

選択範囲ベースとバインドベースのどちらのデータ アクセスでも、ドキュメント コンテンツは、サポートされているすべての Office アプリケーション間で共通のデータ型を通じて公開されます。 2013 Officeでは、3 つの主要なデータ型がサポートされています。

|**データ型**|**説明**|**ホスト アプリケーションのサポート**|
|:-----|:-----|:-----|
|テキスト|選択範囲またはバインド内のデータの文字列表現を提供します。|Excel 2013、Project 2013、および PowerPoint 2013 は、プレーンテキストのみがサポートされます。Word 2013 では、3 つのテキスト形式 (プレーン テキスト、HTML、および Office Open XML (OOXML)) がサポートされます。Excel のセル内でテキストが選択されていると (セル内でテキストの一部のみが選択されている場合でも)、選択範囲ベースのメソッドは、セルのコンテンツ全体の読み取りおよび書き込みを行います。Word および PowerPoint でテキストが選択されていると、選択範囲ベースのメソッドは、選択されている文字の並びのみの読み取りおよび書き込みを行います。Project 2013 および PowerPoint 2013 は、選択範囲ベースのデータ アクセスのみをサポートします。|
|マトリックス|選択範囲またはバインドに含まれるデータを 2 次元の **Array** として提供します (JavaScript で配列の配列として実装されているものです)。たとえば、2 つの列にある 2 つ行の **string** 値は ` [['a', 'b'], ['c', 'd']]` になり、3 つの行を持つ 1 つの列は `[['a'], ['b'], ['c']]` になります。|マトリックス データ アクセスは Excel 2013 および Word 2013 でのみサポートされています。|
|テーブル|選択範囲またはバインド内のデータを [TableData](/javascript/api/office/office.tabledata) オブジェクトとして提供します。 オブジェクト `TableData` は、and プロパティを使用してデータを `headers` 公開 `rows` します。|テーブル データ アクセスは Excel 2013 および Word 2013 でのみサポートされています。|

#### <a name="data-type-coercion"></a>データ型の強制型変換

`Document` and [Binding](/javascript/api/office/office.binding) オブジェクトのデータ アクセス メソッドでは、これらのメソッドの _coercionType_ パラメーターと対応する [CoercionType](/javascript/api/office/office.coerciontype) 列挙値を使用して、目的のデータ型を指定できます。 バインドの実際の形状にかかわらず、さまざまな Office アプリケーションでは、要求されるデータ型にデータを強制的に型変換することによって、共通のデータ型をサポートします。 たとえば、Word の表または段落が選択されている場合、開発者はそれをプレーン テキスト、HTML、Office Open XML、または表として読み取ることを指定でき、API 実装によって必要な変換やデータ変換が行われます。

> [!TIP]
> **データ アクセスにマトリックスを使用する場合と、テーブルの coercionType を使用する場合。** 行と列を追加するときに表形式のデータを動的に拡大する必要がある場合に、テーブル ヘッダーを使用する必要がある場合は、テーブル データ型を使用する必要があります (またはオブジェクト データ アクセスメソッドの _coercionType_ `Document` `Binding` `"table"` `Office.CoercionType.Table`パラメーターを指定します)。 データ構造内の行と列の追加は、テーブルとマトリックス データの両方でサポートされますが、行と列の追加はテーブル データでのみサポートされます。 行と列の追加を計画していない場合に、データにヘッダー機能が必要ない場合は、データ アクセス メソッドの  _coercionType_ `"matrix"` `Office.CoercionType.Matrix`パラメーターを指定して、データを操作するより簡単なモデルを提供するマトリックス データ型を使用する必要があります。

指定された型にデータを強制的に型変換できない場合は、コールバック内の [AsyncResult.status](/javascript/api/office/office.asyncresult#office-office-asyncresult-status-member) プロパティが `"failed"` を返すため、[AsyncResult.error](/javascript/api/office/office.asyncresult#office-office-asyncresult-error-member) プロパティを使用して [Error](/javascript/api/office/office.error) オブジェクトにアクセスし、メソッド呼び出しが失敗した理由を確認できます。

## <a name="work-with-selections-using-the-document-object"></a>Document オブジェクトを使用して選択範囲を使用する

オブジェクト `Document` は、ユーザーの現在の選択範囲に対する読み取りおよび書き込みを "get and set" 方式で行うメソッドを公開します。 これを行うには、オブジェクト `Document` は and メソッド `getSelectedDataAsync` を `setSelectedDataAsync` 提供します。

選択範囲に関する操作の実行方法を示すコード例については、「[ドキュメントまたはスプレッドシート内のアクティブな選択範囲へのデータの読み取りおよび書き込み](read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md)」を参照してください。

## <a name="work-with-bindings-using-the-bindings-and-binding-objects"></a>Bindings オブジェクトと Binding オブジェクトを使用したバインドの操作

バインドベースのデータ アクセスを使用すると、コンテンツ アドインおよび作業ウィンドウ アドインで、バインドに関連付けられた識別子を介して、ドキュメントまたはスプレッドシートの特定の領域に一貫性のあるアクセスが可能になります。 アドインは、最初に、ドキュメントの部分と一意の ID を関連付けるメソッドのいずれか ([addFromPromptAsync](/javascript/api/office/office.bindings#office-office-bindings-addfrompromptasync-member(1))、[addFromSelectionAsync](/javascript/api/office/office.bindings#office-office-bindings-addfromselectionasync-member(1))、または [addFromNamedItemAsync](/javascript/api/office/office.bindings#office-office-bindings-addfromnameditemasync-member(1))) を呼び出すことによって、バインドを確立する必要があります。 バインドが確立されると、アドインは提供された ID を使用して、ドキュメントまたはスプレッドシート内の関連付けられた領域に含まれるデータにアクセスできます。 バインドを作成すると、アドインに次の値が提供されます。

- 表、範囲、またはテキスト (隣接する一連の文字) など、サポートされている Office アプリケーション全体に共通のデータ構造へのアクセスを許可します。

- ユーザーによる選択を必要とせずに、読み取り/書き込み操作ができます。

- アドインとドキュメント内のデータの間にリレーションシップが確立されます。バインドはドキュメント内に保持され、後でアクセスできます。

また、バインドを確立すると、ドキュメントまたはスプレッドシートの特定の領域を範囲とする、データおよび選択範囲の変更イベントをサブスクライブできます。つまり、ドキュメントまたはスプレッドシート全体の全般的な変更ではなく、バインドされた領域内で発生する変更のみがアドインに通知されます。

[Bindings](/javascript/api/office/office.bindings) オブジェクトが公開している [getAllAsync](/javascript/api/office/office.bindings#office-office-bindings-getallasync-member(1)) メソッドを使用すると、ドキュメントまたはスプレッドシートで確立されている一連のすべてのバインドにアクセスできます。 個々のバインドに ID でアクセスするには、[Bindings.getBindingByIdAsync](/javascript/api/office/office.bindings#office-office-bindings-getbyidasync-member(1)) メソッドまたは [Office.select](/javascript/api/office) メソッドを使用します。 `Bindings` [addFromSelectionAsync](/javascript/api/office/office.bindings#office-office-bindings-addfromselectionasync-member(1))、[addFromPromptAsync](/javascript/api/office/office.bindings#office-office-bindings-addfrompromptasync-member(1))、[addFromNamedItemAsync](/javascript/api/office/office.bindings#office-office-bindings-addfromnameditemasync-member(1))、[または releaseByIdAsync](/javascript/api/office/office.bindings#office-office-bindings-releasebyidasync-member(1)) のいずれかのメソッドを使用して、新しいバインドを確立し、既存のバインドを削除できます。

バインドを作成するときに  _bindingType_ `addFromSelectionAsync`パラメーターを使用して指定するバインドには、3 つの異なる種類 `addFromPromptAsync` `addFromNamedItemAsync` があります。

|**バインドの種類**|**説明**|**ホスト アプリケーションのサポート**|
|:-----|:-----|:-----|
|テキスト バインド|テキストとして表現できるドキュメントの領域にバインドします。|Word では、連続する選択範囲の大部分が有効ですが、Excel では、単一セルの範囲のみがテキスト バインドの対象です。Excel では、プレーン テキストのみがサポートされます。Word では、3 つの形式 (プレーン テキスト、HTML、および Open XML for Office) がサポートされます。|
|マトリックス バインド|ヘッダーがない表形式のデータが含まれるドキュメントの固定領域にバインドします。マトリックス バインド内のデータは、2 次元の **Array** として書き込みまたは読み取りが行われます。JavaScript では、これは、配列の配列として実装されています。たとえば、2 列の **string** 値が 2 行ある場合は ` [['a', 'b'], ['c', 'd']]` のように書き込みまたは読み取りが行われ、1 列が 3 行ある場合は `[['a'], ['b'], ['c']]` のように書き込みまたは読み取りが行われます。|Excel では、セルの連続する選択範囲を使用してマトリックス バインドを確立できます。Word では、表のみがマトリックス バインドをサポートします。|
|テーブル バインド|ヘッダーがある表が含まれるドキュメントの領域にバインドします。 テーブル バインド内のデータは、[TableData](/javascript/api/office/office.tabledata) オブジェクトとして書き込みまたは読み取りが行われます。 オブジェクト `TableData` は、ヘッダーと行のプロパティを使用 **してデータ** を **公開** します。|Excel または Word の表はすべて、テーブル バインドの基礎にできます。テーブル バインドを確立すると、ユーザーが表に追加する新しい各行または各列が、自動的にバインドに含まれます。 |

<br/>

オブジェクトの 3 つの "add" メソッドのいずれかを使用してバインドを作成した後、対応するオブジェクトの[メソッド (MatrixBinding](/javascript/api/office/office.matrixbinding)、[TableBinding](/javascript/api/office/office.tablebinding)、[または TextBinding](/javascript/api/office/office.textbinding)) `Bindings` を使用して、バインドのデータとプロパティを処理できます。 これら 3 つのオブジェクトはすべて、バインドされたデータを操作できるオブジェクトの [getDataAsync](/javascript/api/office/office.binding#office-office-binding-getdataasync-member(1)) メソッドと [setDataAsync](/javascript/api/office/office.binding#office-office-binding-setdataasync-member(1)) `Binding` メソッドを継承します。

バインドに関する操作の実行方法を示すコード例については、「[ドキュメントまたはスプレッドシート内の領域へのバインド](bind-to-regions-in-a-document-or-spreadsheet.md)」を参照してください。

## <a name="work-with-custom-xml-parts-using-the-customxmlparts-and-customxmlpart-objects"></a>CustomXmlParts オブジェクトと CustomXmlPart オブジェクトを使用してカスタム XML パーツを操作する

 **適用対象:** Word の作業ウィンドウ アドイン

API の [CustomXmlParts](/javascript/api/office/office.customxmlparts) オブジェクトと [CustomXmlPart](/javascript/api/office/office.customxmlpart) オブジェクトを使用すると、Word 文書内のカスタム XML パーツにアクセスできます。これにより、文書のコンテンツに対する XML 主導の操作が可能になります。 and オブジェクトの操作のデモンストレーション`CustomXmlParts``CustomXmlPart`については、[Word-add-in-Work-with-custom-XML-parts](https://github.com/OfficeDev/Word-Add-in-Work-with-custom-XML-parts) コード サンプルを参照してください。

## <a name="work-with-the-entire-document-using-the-getfileasync-method"></a>getFileAsync メソッドを使用してドキュメント全体を操作する

 **適用対象:** Word および PowerPoint の作業ウィンドウ アドイン

[Document.getFileAsync](/javascript/api/office/office.document#office-office-document-getfileasync-member(1)) メソッド、および [File](/javascript/api/office/office.file) オブジェクトと [Slice](/javascript/api/office/office.slice) オブジェクトのメンバーは、一度に最大で 4 MB ずつのスライス (チャンク) に分割して Word および PowerPoint ドキュメント ファイル全体を取得する機能を提供します。詳細については、「[PowerPoint または Word 用アドインからドキュメント全体を取得する](../word/get-the-whole-document-from-an-add-in-for-word.md)」を参照してください。

## <a name="mailbox-object"></a>Mailbox オブジェクト

**適用対象:** Outlook アドイン

Outlook アドインでは、主に [Mailbox](/javascript/api/outlook/office.mailbox) オブジェクトにより公開されている API のサブセットを使用します。Outlook アドイン専用のオブジェクトおよびメンバー (たとえば、[Item](/javascript/api/outlook/office.item) オブジェクトなど) にアクセスするには、次のコード行に示すように、[Context](/javascript/api/office/office.context#office-office-context-mailbox-member) オブジェクトの **mailbox** プロパティを使用して、**Mailbox** オブジェクトにアクセスします。

```js
// Access the Item object.
var item = Office.context.mailbox.item;

```

さらに、Outlookは次のオブジェクトを使用できます。

- `Office` object: 初期化用。

- `Context` object: コンテンツにアクセスし、言語プロパティを表示します。

- `RoamingSettings`object: アドインOutlookカスタム設定を、アドインがインストールされているユーザーのメールボックスに保存します。

Outlook アドインでの JavaScript の使用については、「[Outlook アドイン](../outlook/outlook-add-ins-overview.md)」を参照してください。
