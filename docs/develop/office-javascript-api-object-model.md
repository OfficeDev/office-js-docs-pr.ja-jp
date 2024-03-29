---
title: 共通 JavaScript API オブジェクト モデル
description: Office JavaScript の一般的な API オブジェクト モデルについて説明します。
ms.date: 07/07/2022
ms.localizationpriority: medium
ms.openlocfilehash: 1b856866c903a61a04bcbb232790649147fdb7fc
ms.sourcegitcommit: b6a3815a1ad17f3522ca35247a3fd5d7105e174e
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/22/2022
ms.locfileid: "66958623"
---
# <a name="common-javascript-api-object-model"></a>共通 JavaScript API オブジェクト モデル

[!include[information about the common API](../includes/alert-common-api-info.md)]

Office JavaScript API を使用すると、Office クライアント アプリケーションの基になる機能にアクセスできます。 このアクセスの大部分はいくつかの重要なオブジェクトを通過します。 [Context](#context-object) オブジェクトによって、初期化した後、ランタイム環境にアクセスできるようになります。 [Document](#document-object) オブジェクトによって、Excel、PowerPoint、Word ドキュメントを操作する許可が与えられます。 [Mailbox](#mailbox-object) オブジェクトを使用すると、Outlook アドインからメッセージ、予定、ユーザー プロファイルにアクセスできます。 これらの高度なオブジェクト間のリレーションシップを理解することは、Office アドインの基礎です。

## <a name="context-object"></a>Context オブジェクト

**適用対象:** すべてのアドインの種類

アドインが[初期化](initialize-add-in.md)されると、ランタイム環境でやり取りできるさまざまなオブジェクトが多数あります。 アドインのランタイム コンテキストは [Context](/javascript/api/office/office.context) オブジェクトによって API で反映されます。 **Context** は、[Document](/javascript/api/office/office.document) オブジェクトや [Mailbox](/javascript/api/outlook/office.mailbox) オブジェクトなど、API の最重要オブジェクトにアクセスできるメイン オブジェクトです。最重要オブジェクトはさらにドキュメントやメールボックスのコンテンツにアクセスできます。

たとえば、作業ウィンドウ アドインまたはコンテンツ アドインにおいて、[Context](/javascript/api/office/office.context#office-office-context-document-member) オブジェクトの **document** プロパティを使用して、**Document** オブジェクトのプロパティおよびメソッドにアクセスし、Word 文書、Excel ワークシート、または Project スケジュールのコンテンツとやり取りできます。同様に、Outlook アドインにおいて、[Context](/javascript/api/office/office.context#office-office-context-mailbox-member) オブジェクトの **mailbox** プロパティを使用して、**Mailbox** オブジェクトのプロパティおよびメソッドにアクセスし、メッセージ、会議出席依頼または予定のコンテンツとやり取りできます。

**Context** オブジェクトは [contentLanguage](/javascript/api/office/office.context#office-office-context-contentlanguage-member) プロパティと [displayLanguage](/javascript/api/office/office.context#office-office-context-displaylanguage-member) プロパティへのアクセスも提供します。これにより、ドキュメントまたはアイテムまたは Office アプリケーションで使用されるロケール (言語) を決定できます。 [roamingSettings](/javascript/api/office/office.context#office-office-context-roamingsettings-member) プロパティによって、[RoamingSettings](/javascript/api/office/office.context#office-office-context-roamingsettings-member) オブジェクトのメンバーにアクセスできます。このオブジェクトによって、個々のユーザーのメールボックスに対してアドインに固有の設定が保存されます。 最後に、**Context** オブジェクトの [ui](/javascript/api/office/office.context#office-office-context-ui-member) プロパティを使用すると、アドインでポップアップ ダイアログを開始できます。

## <a name="document-object"></a>Document オブジェクト

**適用対象:** コンテンツ アドインおよび作業ウィンドウ アドインの種類

Excel、PowerPoint、および Word のドキュメント データを操作するために、API には [Document](/javascript/api/office/office.document) オブジェクトが用意されています。 オブジェクト メンバーを使用 `Document` して、次の方法でデータにアクセスできます。

- テキスト、隣接するセル (マトリックス)、またはテーブルの形式のアクティブな選択範囲への読み取りと書き込み。

- 表形式のデータ (マトリックスまたはテーブル)。

- バインディング (オブジェクトの "add" メソッドを使用して `Bindings` 作成)。

- カスタム XML パーツ (Word の場合のみ)。

- ドキュメント上のアドインごとに保持する設定またはアドインの状態。

オブジェクトを `Document` 使用して、Project ドキュメント内のデータを操作することもできます。 API の Project 固有の機能については、[ProjectDocument](/javascript/api/office/office.document) 抽象クラスのメンバー内に説明文があります。 Project 用の作業ウィンドウ アドインの作成の詳細については、「[Project 用の作業ウィンドウ アドイン](../project/project-add-ins.md)」を参照してください。

これらの形式のデータ アクセスはすべて、抽象 `Document` オブジェクトのインスタンスから開始されます。

作業ウィンドウまたはコンテンツ アドインがオブジェクトの`Document`[ドキュメント](/javascript/api/office/office.context#office-office-context-document-member) プロパティを使用して初期化されるときに、オブジェクトのインスタンスに`Context`アクセスできます。 このオブジェクトは `Document` 、Word および Excel ドキュメント間で共有される一般的なデータ アクセス方法を定義し、Word ドキュメントのオブジェクトへの `CustomXmlParts` アクセスも提供します。

このオブジェクトは `Document` 、開発者がドキュメントの内容にアクセスするための 4 つの方法をサポートしています。

- 選択範囲ベースのアクセス

- バインドベースのアクセス

- カスタム XML パーツベースのアクセス (Word の場合のみ)

- ドキュメント全体へのアクセス (PowerPoint および Word のみ)

選択範囲ベースおよびバインドベースのデータ アクセス方法のしくみを理解するために、まず、データ アクセス API が、異なる Office アプリケーション間で一貫性のあるデータ アクセスを提供する方法について説明します。

### <a name="consistent-data-access-across-office-applications"></a>Office アプリケーション間での一貫性のあるデータ アクセス

 **適用対象:** コンテンツ アドインおよび作業ウィンドウ アドインの種類

Office JavaScript API は、さまざまな Office ドキュメント間でシームレスに機能する拡張機能を作成するために、一般的なデータ型と、異なるドキュメントコンテンツを 3 つの一般的なデータ型に強制する機能を通じて、各 Office アプリケーションの詳細を抽象化します。

#### <a name="common-data-types"></a>共通のデータ型

選択範囲ベースとバインドベースのどちらのデータ アクセスでも、ドキュメント コンテンツは、サポートされているすべての Office アプリケーション間で共通のデータ型を通じて公開されます。 Office 2013 では、3 つの主要なデータ型がサポートされています。

|**データ型**|**説明**|**ホスト アプリケーションのサポート**|
|:-----|:-----|:-----|
|テキスト|選択範囲またはバインド内のデータの文字列表現を提供します。|Excel 2013、Project 2013、および PowerPoint 2013 は、プレーンテキストのみがサポートされます。Word 2013 では、3 つのテキスト形式 (プレーン テキスト、HTML、および Office Open XML (OOXML)) がサポートされます。Excel のセル内でテキストが選択されていると (セル内でテキストの一部のみが選択されている場合でも)、選択範囲ベースのメソッドは、セルのコンテンツ全体の読み取りおよび書き込みを行います。Word および PowerPoint でテキストが選択されていると、選択範囲ベースのメソッドは、選択されている文字の並びのみの読み取りおよび書き込みを行います。Project 2013 および PowerPoint 2013 は、選択範囲ベースのデータ アクセスのみをサポートします。|
|マトリックス|選択範囲またはバインドに含まれるデータを 2 次元の **Array** として提供します (JavaScript で配列の配列として実装されているものです)。たとえば、2 つの列にある 2 つ行の **string** 値は ` [['a', 'b'], ['c', 'd']]` になり、3 つの行を持つ 1 つの列は `[['a'], ['b'], ['c']]` になります。|マトリックス データ アクセスは Excel 2013 および Word 2013 でのみサポートされています。|
|テーブル|選択範囲またはバインド内のデータを [TableData](/javascript/api/office/office.tabledata) オブジェクトとして提供します。 オブジェクトは`TableData`、プロパティを`rows`使用してデータを公開します`headers`。|テーブル データ アクセスは Excel 2013 および Word 2013 でのみサポートされています。|

#### <a name="data-type-coercion"></a>データ型の強制型変換

および [Binding](/javascript/api/office/office.binding) オブジェクトの`Document`データ アクセス メソッドでは、これらのメソッドの _coercionType_ パラメーターと対応する [CoercionType](/javascript/api/office/office.coerciontype) 列挙値を使用して、目的のデータ型を指定できます。 バインドの実際の形状にかかわらず、さまざまな Office アプリケーションでは、要求されるデータ型にデータを強制的に型変換することによって、共通のデータ型をサポートします。 たとえば、Word の表または段落が選択されている場合、開発者はそれをプレーン テキスト、HTML、Office Open XML、または表として読み取ることを指定でき、API 実装によって必要な変換やデータ変換が行われます。

> [!TIP]
> **データ アクセスにマトリックスを使用する場合と、テーブルの coercionType を使用する場合。** 行と列が追加されたときにテーブル データを動的に拡張する必要があり、テーブル ヘッダーを操作する必要がある場合は、テーブル データ型を使用する必要があります (またはオブジェクト データ アクセス方法`"table"`の `Document` _coercionType_ パラメーターを指定するか、または`Office.CoercionType.Table``Binding`指定します)。 データ構造内の行と列の追加は、テーブルとマトリックス データの両方でサポートされますが、行と列の追加はテーブル データでのみサポートされます。 行と列を追加する予定がなく、データにヘッダー機能が必要ない場合は、マトリックス データ型を使用する必要があります (データ アクセス方法`"matrix"`の _coercionType_ パラメーターを指定するか、または`Office.CoercionType.Matrix`データと対話するより簡単なモデルを提供します)。

指定された型にデータを強制的に型変換できない場合は、コールバック内の [AsyncResult.status](/javascript/api/office/office.asyncresult#office-office-asyncresult-status-member) プロパティが `"failed"` を返すため、[AsyncResult.error](/javascript/api/office/office.asyncresult#office-office-asyncresult-error-member) プロパティを使用して [Error](/javascript/api/office/office.error) オブジェクトにアクセスし、メソッド呼び出しが失敗した理由を確認できます。

## <a name="work-with-selections-using-the-document-object"></a>Document オブジェクトを使用して選択範囲を操作する

このオブジェクトは `Document` 、ユーザーの現在の選択内容を "get and set" の方法で読み取りおよび書き込みできるメソッドを公開します。 これを行うために、オブジェクトは`Document`メソッドと`setSelectedDataAsync`メソッドを提供します`getSelectedDataAsync`。

選択範囲に関する操作の実行方法を示すコード例については、「[ドキュメントまたはスプレッドシート内のアクティブな選択範囲へのデータの読み取りおよび書き込み](read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md)」を参照してください。

## <a name="work-with-bindings-using-the-bindings-and-binding-objects"></a>Bindings オブジェクトと Binding オブジェクトを使用してバインドを操作する

バインドベースのデータ アクセスを使用すると、コンテンツ アドインおよび作業ウィンドウ アドインで、バインドに関連付けられた識別子を介して、ドキュメントまたはスプレッドシートの特定の領域に一貫性のあるアクセスが可能になります。 アドインは、最初に、ドキュメントの部分と一意の ID を関連付けるメソッドのいずれか ([addFromPromptAsync](/javascript/api/office/office.bindings#office-office-bindings-addfrompromptasync-member(1))、[addFromSelectionAsync](/javascript/api/office/office.bindings#office-office-bindings-addfromselectionasync-member(1))、または [addFromNamedItemAsync](/javascript/api/office/office.bindings#office-office-bindings-addfromnameditemasync-member(1))) を呼び出すことによって、バインドを確立する必要があります。 バインドが確立されると、アドインは提供された ID を使用して、ドキュメントまたはスプレッドシート内の関連付けられた領域に含まれるデータにアクセスできます。 バインドを作成すると、アドインに次の値が提供されます。

- 表、範囲、またはテキスト (隣接する一連の文字) など、サポートされている Office アプリケーション全体に共通のデータ構造へのアクセスを許可します。

- ユーザーによる選択を必要とせずに、読み取り/書き込み操作ができます。

- アドインとドキュメント内のデータの間にリレーションシップが確立されます。バインドはドキュメント内に保持され、後でアクセスできます。

また、バインドを確立すると、ドキュメントまたはスプレッドシートの特定の領域を範囲とする、データおよび選択範囲の変更イベントをサブスクライブできます。つまり、ドキュメントまたはスプレッドシート全体の全般的な変更ではなく、バインドされた領域内で発生する変更のみがアドインに通知されます。

[Bindings](/javascript/api/office/office.bindings) オブジェクトが公開している [getAllAsync](/javascript/api/office/office.bindings#office-office-bindings-getallasync-member(1)) メソッドを使用すると、ドキュメントまたはスプレッドシートで確立されている一連のすべてのバインドにアクセスできます。 個々のバインドには、 [Bindings.getBindingByIdAsync](/javascript/api/office/office.bindings#office-office-bindings-getbyidasync-member(1)) メソッドまたは [Office.select](/javascript/api/office) 関数のいずれかを使用して、ID でアクセスできます。 [addFromSelectionAsync、addFromPromptAsync](/javascript/api/office/office.bindings#office-office-bindings-addfromselectionasync-member(1))、[addFromNamedItemAsync](/javascript/api/office/office.bindings#office-office-bindings-addfromnameditemasync-member(1))、[releaseByIdAsync](/javascript/api/office/office.bindings#office-office-bindings-releasebyidasync-member(1)) のいずれかのメソッドを使用[](/javascript/api/office/office.bindings#office-office-bindings-addfrompromptasync-member(1))して、新しいバインドを確立したり、既存の`Bindings`バインドを削除したりすることもできます。

バインドを`addFromNamedItemAsync`作成するときに _bindingType_ パラメーターで指定するバインド`addFromSelectionAsync``addFromPromptAsync`には、次の 3 種類があります。

|**バインドの種類**|**説明**|**ホスト アプリケーションのサポート**|
|:-----|:-----|:-----|
|テキスト バインド|テキストとして表現できるドキュメントの領域にバインドします。|Word では、連続する選択範囲の大部分が有効ですが、Excel では、単一セルの範囲のみがテキスト バインドの対象です。Excel では、プレーン テキストのみがサポートされます。Word では、3 つの形式 (プレーン テキスト、HTML、および Open XML for Office) がサポートされます。|
|マトリックス バインド|ヘッダーがない表形式のデータが含まれるドキュメントの固定領域にバインドします。マトリックス バインド内のデータは、2 次元の **Array** として書き込みまたは読み取りが行われます。JavaScript では、これは、配列の配列として実装されています。たとえば、2 列の **string** 値が 2 行ある場合は ` [['a', 'b'], ['c', 'd']]` のように書き込みまたは読み取りが行われ、1 列が 3 行ある場合は `[['a'], ['b'], ['c']]` のように書き込みまたは読み取りが行われます。|Excel では、セルの連続する選択範囲を使用してマトリックス バインドを確立できます。Word では、表のみがマトリックス バインドをサポートします。|
|テーブル バインド|ヘッダーがある表が含まれるドキュメントの領域にバインドします。 テーブル バインド内のデータは、[TableData](/javascript/api/office/office.tabledata) オブジェクトとして書き込みまたは読み取りが行われます。 オブジェクトは `TableData` 、 **ヘッダー** と **行** のプロパティを使用してデータを公開します。|Excel または Word の表はすべて、テーブル バインドの基礎にできます。テーブル バインドを確立すると、ユーザーが表に追加する新しい各行または各列が、自動的にバインドに含まれます。 |

<br/>

オブジェクトの 3 つの "add" メソッド `Bindings` のいずれかを使用してバインドを作成した後は、対応するオブジェクト ( [MatrixBinding、TableBinding](/javascript/api/office/office.matrixbinding)、 [TextBinding](/javascript/api/office/office.tablebinding)) のメソッドを使用して、バインドのデータとプロパティ [を](/javascript/api/office/office.textbinding)操作できます。 これら 3 つのオブジェクトはすべて、バインドされたデータを操作できるオブジェクトの `Binding` [getDataAsync](/javascript/api/office/office.binding#office-office-binding-getdataasync-member(1)) メソッドと [setDataAsync](/javascript/api/office/office.binding#office-office-binding-setdataasync-member(1)) メソッドを継承します。

バインドに関する操作の実行方法を示すコード例については、「[ドキュメントまたはスプレッドシート内の領域へのバインド](bind-to-regions-in-a-document-or-spreadsheet.md)」を参照してください。

## <a name="work-with-custom-xml-parts-using-the-customxmlparts-and-customxmlpart-objects"></a>CustomXmlParts オブジェクトと CustomXmlPart オブジェクトを使用してカスタム XML パーツを操作する

 **適用対象:** Word の作業ウィンドウ アドイン

API の [CustomXmlParts](/javascript/api/office/office.customxmlparts) オブジェクトと [CustomXmlPart](/javascript/api/office/office.customxmlpart) オブジェクトを使用すると、Word 文書内のカスタム XML パーツにアクセスできます。これにより、文書のコンテンツに対する XML 主導の操作が可能になります。 オブジェクトの操作`CustomXmlParts`のデモンストレーションについては、[Word-add-in-Work-with-custom-XML-parts](https://github.com/OfficeDev/Word-Add-in-Work-with-custom-XML-parts) コード サンプルを参照`CustomXmlPart`してください。

## <a name="work-with-the-entire-document-using-the-getfileasync-method"></a>getFileAsync メソッドを使用してドキュメント全体を操作する

 **適用対象:** Word および PowerPoint の作業ウィンドウ アドイン

[Document.getFileAsync](/javascript/api/office/office.document#office-office-document-getfileasync-member(1)) メソッド、および [File](/javascript/api/office/office.file) オブジェクトと [Slice](/javascript/api/office/office.slice) オブジェクトのメンバーは、一度に最大で 4 MB ずつのスライス (チャンク) に分割して Word および PowerPoint ドキュメント ファイル全体を取得する機能を提供します。詳細については、「[PowerPoint または Word 用アドインからドキュメント全体を取得する](../word/get-the-whole-document-from-an-add-in-for-word.md)」を参照してください。

## <a name="mailbox-object"></a>Mailbox オブジェクト

**適用対象:** Outlook アドイン

[!INCLUDE [Mailbox object information](../includes/mailbox-object-desc.md)]
