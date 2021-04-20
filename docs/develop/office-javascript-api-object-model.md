---
title: 共通 JavaScript API オブジェクト モデル
description: Office JavaScript 共通 API オブジェクトモデルについて
ms.date: 04/30/2020
localization_priority: Normal
ms.openlocfilehash: 37d2bca0aa4aadfc6ab7ef00d76d74e9acde4711
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/28/2020
ms.locfileid: "47293255"
---
# <a name="common-javascript-api-object-model"></a>共通 JavaScript API オブジェクト モデル

[!include[information about the common API](../includes/alert-common-api-info.md)]

Office JavaScript Api は、Office クライアントアプリケーションの基礎となる機能にアクセスできるようにします。 このアクセスの大部分はいくつかの重要なオブジェクトを通過します。 [Context](#context-object) オブジェクトによって、初期化した後、ランタイム環境にアクセスできるようになります。 [Document](#document-object) オブジェクトによって、Excel、PowerPoint、Word ドキュメントを操作する許可が与えられます。 [メールボックス](#mailbox-object)オブジェクトは、メッセージ、予定、およびユーザープロファイルへの Outlook アドインのアクセスを提供します。 これらの高レベルオブジェクト間の関係を理解することは、Office アドインの基礎となります。

## <a name="context-object"></a>Context オブジェクト

**適用対象:** すべてのアドインの種類

アドインが[初期化](initialize-add-in.md)されると、ランタイム環境でやり取りできるさまざまなオブジェクトが多数あります。 アドインのランタイム コンテキストは [Context](/javascript/api/office/office.context) オブジェクトによって API で反映されます。 **Context** は、[Document](/javascript/api/office/office.document) オブジェクトや [Mailbox](/javascript/api/outlook/office.mailbox) オブジェクトなど、API の最重要オブジェクトにアクセスできるメイン オブジェクトです。最重要オブジェクトはさらにドキュメントやメールボックスのコンテンツにアクセスできます。

たとえば、作業ウィンドウ アドインまたはコンテンツ アドインにおいて、[Context](/javascript/api/office/office.context#document) オブジェクトの **document** プロパティを使用して、**Document** オブジェクトのプロパティおよびメソッドにアクセスし、Word 文書、Excel ワークシート、または Project スケジュールのコンテンツとやり取りできます。同様に、Outlook アドインにおいて、[Context](/javascript/api/office/office.context#mailbox) オブジェクトの **mailbox** プロパティを使用して、**Mailbox** オブジェクトのプロパティおよびメソッドにアクセスし、メッセージ、会議出席依頼または予定のコンテンツとやり取りできます。

また、 **Context** オブジェクトでは、ドキュメント、アイテム、または Office アプリケーションで使用されるロケール (言語) を決定できるように、 [Contentlanguage](/javascript/api/office/office.context#contentlanguage) プロパティと [displaylanguage](/javascript/api/office/office.context#displaylanguage) プロパティへのアクセスも提供します。 [roamingSettings](/javascript/api/office/office.context#roamingsettings) プロパティによって、[RoamingSettings](/javascript/api/office/office.context#roamingsettings) オブジェクトのメンバーにアクセスできます。このオブジェクトによって、個々のユーザーのメールボックスに対してアドインに固有の設定が保存されます。 最後に、**Context** オブジェクトの [ui](/javascript/api/office/office.context#ui) プロパティを使用すると、アドインでポップアップ ダイアログを開始できます。


## <a name="document-object"></a>Document オブジェクト

**適用対象:** コンテンツ アドインおよび作業ウィンドウ アドインの種類

Excel、PowerPoint、および Word のドキュメント データを操作するために、API には [Document](/javascript/api/office/office.document) オブジェクトが用意されています。 `Document`オブジェクトのメンバーを使用して、次の方法でデータにアクセスできます。

- テキスト、隣接するセル (マトリックス)、またはテーブルの形式のアクティブな選択範囲への読み取りと書き込み。

- 表形式のデータ (マトリックスまたはテーブル)。

- バインド (オブジェクトの "add" メソッドで作成 `Bindings` )。

- カスタム XML パーツ (Word の場合のみ)。

- ドキュメント上のアドインごとに保持する設定またはアドインの状態。

オブジェクトを使用し `Document` て、プロジェクトドキュメント内のデータを操作することもできます。 API の Project 固有の機能については、[ProjectDocument](/javascript/api/office/office.document) 抽象クラスのメンバー内に説明文があります。 Project 用の作業ウィンドウ アドインの作成の詳細については、「[Project 用の作業ウィンドウ アドイン](../project/project-add-ins.md)」を参照してください。

これらのデータアクセスの形式はすべて、abstract オブジェクトのインスタンスから開始 `Document` します。

オブジェクトの `Document` [document](/javascript/api/office/office.context#document) プロパティを使用して、作業ウィンドウアドインまたはコンテンツアドインを初期化したときに、オブジェクトのインスタンスにアクセスでき `Context` ます。 オブジェクトは、 `Document` word 文書と Excel 文書間で共有される共通のデータアクセス関数を定義し、word 文書のオブジェクトへのアクセスも提供し `CustomXmlParts` ます。

この `Document` オブジェクトは、開発者がドキュメントコンテンツにアクセスするための4つの方法をサポートしています。


- 選択範囲ベースのアクセス

- バインドベースのアクセス

- カスタム XML パーツベースのアクセス (Word の場合のみ)

- ドキュメント全体へのアクセス (PowerPoint および Word のみ)

選択範囲ベースおよびバインドベースのデータ アクセス方法のしくみを理解するために、まず、データ アクセス API が、異なる Office アプリケーション間で一貫性のあるデータ アクセスを提供する方法について説明します。


### <a name="consistent-data-access-across-office-applications"></a>Office アプリケーション間での一貫性のあるデータ アクセス

 **適用対象:** コンテンツ アドインおよび作業ウィンドウ アドインの種類

異なる Office ドキュメント間でシームレスに動作する拡張機能を作成するために、Office JavaScript API は、共通のデータ型と異なるドキュメントのコンテンツを3つの共通のデータ型に強制的に変換する機能を通じて、各 Office アプリケーションの特殊性を抽象化します。


#### <a name="common-data-types"></a>共通のデータ型

選択範囲ベースとバインドベースのどちらのデータ アクセスでも、ドキュメント コンテンツは、サポートされているすべての Office アプリケーション間で共通のデータ型を通じて公開されます。Office 2013 では、3 つの主要なデータ型がサポートされています。



|**データ型**|**説明**|**ホスト アプリケーションのサポート**|
|:-----|:-----|:-----|
|テキスト|選択範囲またはバインド内のデータの文字列表現を提供します。|Excel 2013、Project 2013、および PowerPoint 2013 は、プレーンテキストのみがサポートされます。Word 2013 では、3 つのテキスト形式 (プレーン テキスト、HTML、および Office Open XML (OOXML)) がサポートされます。Excel のセル内でテキストが選択されていると (セル内でテキストの一部のみが選択されている場合でも)、選択範囲ベースのメソッドは、セルのコンテンツ全体の読み取りおよび書き込みを行います。Word および PowerPoint でテキストが選択されていると、選択範囲ベースのメソッドは、選択されている文字の並びのみの読み取りおよび書き込みを行います。Project 2013 および PowerPoint 2013 は、選択範囲ベースのデータ アクセスのみをサポートします。|
|マトリックス|選択範囲またはバインドに含まれるデータを 2 次元の **Array** として提供します (JavaScript で配列の配列として実装されているものです)。たとえば、2 つの列にある 2 つ行の **string** 値は ` [['a', 'b'], ['c', 'd']]` になり、3 つの行を持つ 1 つの列は `[['a'], ['b'], ['c']]` になります。|マトリックス データ アクセスは Excel 2013 および Word 2013 でのみサポートされています。|
|テーブル|選択範囲またはバインド内のデータを [TableData](/javascript/api/office/office.tabledata) オブジェクトとして提供します。 `TableData`オブジェクトは、プロパティとプロパティを通じてデータを公開し `headers` `rows` ます。|テーブル データ アクセスは Excel 2013 および Word 2013 でのみサポートされています。|

#### <a name="data-type-coercion"></a>データ型の強制型変換

および Binding オブジェクトのデータアクセス方法は、 `Document` これらのメソッドの_coercionType_パラメーターと対応[Binding](/javascript/api/office/office.binding)する[coercionType](/javascript/api/office/office.coerciontype)列挙値を使用して、目的のデータ型を指定することをサポートしています。 バインドの実際の形状にかかわらず、さまざまな Office アプリケーションでは、要求されるデータ型にデータを強制的に型変換することによって、共通のデータ型をサポートします。 たとえば、Word の表または段落が選択されている場合、開発者はそれをプレーン テキスト、HTML、Office Open XML、または表として読み取ることを指定でき、API 実装によって必要な変換やデータ変換が行われます。


> [!TIP]
> **データ アクセスにマトリックスを使用する場合と、テーブルの coercionType を使用する場合。** 行と列が追加されたときに表形式のデータを動的に拡張する必要があり、テーブルのヘッダーを処理する必要がある場合は、 _coercionType_テーブルのデータ型を使用する必要があり `Document` ます (または、またはオブジェクトデータアクセスメソッドの coercionType パラメーターを指定するか、または `Binding` `"table"` `Office.CoercionType.Table` )。 データ構造内の行と列の追加は、テーブルとマトリックス データの両方でサポートされますが、行と列の追加はテーブル データでのみサポートされます。 行と列の追加を計画しておらず、データにヘッダー機能が必要ない場合は、マトリックスデータ型を使用する必要があります (またはデータアクセス方法の  _coercionType_ パラメーターを指定することによって、データを操作する `"matrix"` `Office.CoercionType.Matrix` ための簡単なモデルが提供されます)。

指定された型にデータを強制的に型変換できない場合は、コールバック内の [AsyncResult.status](/javascript/api/office/office.asyncresult#status) プロパティが `"failed"` を返すため、[AsyncResult.error](/javascript/api/office/office.asyncresult#error) プロパティを使用して [Error](/javascript/api/office/office.error) オブジェクトにアクセスし、メソッド呼び出しが失敗した理由を確認できます。


## <a name="working-with-selections-using-the-document-object"></a>Document オブジェクトによる選択範囲の操作


オブジェクトは、 `Document` "get and set" という方法でユーザーの現在の選択範囲の読み取りと書き込みを行うことができるメソッドを公開します。 そのために、 `Document` オブジェクトは `getSelectedDataAsync` メソッドとメソッドを提供し `setSelectedDataAsync` ます。

選択範囲に関する操作の実行方法を示すコード例については、「[ドキュメントまたはスプレッドシート内のアクティブな選択範囲へのデータの読み取りおよび書き込み](read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md)」を参照してください。


## <a name="working-with-bindings-using-the-bindings-and-binding-objects"></a>Bindings オブジェクトおよび Binding オブジェクトによるバインドの操作


バインドベースのデータ アクセスを使用すると、コンテンツ アドインおよび作業ウィンドウ アドインで、バインドに関連付けられた識別子を介して、ドキュメントまたはスプレッドシートの特定の領域に一貫性のあるアクセスが可能になります。アドインは、最初に、ドキュメントの部分と一意の ID を関連付けるメソッドのいずれか ([addFromPromptAsync](/javascript/api/office/office.bindings#addfrompromptasync-bindingtype--options--callback-)、[addFromSelectionAsync](/javascript/api/office/office.bindings#addfromselectionasync-bindingtype--options--callback-)、または [addFromNamedItemAsync](/javascript/api/office/office.bindings#addfromnameditemasync-itemname--bindingtype--options--callback-)) を呼び出すことによって、バインドを確立する必要があります。バインドが確立されると、アドインは提供された ID を使用して、ドキュメントまたはスプレッドシート内の関連付けられた領域に含まれるデータにアクセスできます。バインドを作成すると、アドインには次のようなメリットがあります。


- 表、範囲、またはテキスト (隣接する一連の文字) など、サポートされている Office アプリケーション全体に共通のデータ構造へのアクセスを許可します。

- ユーザーによる選択を必要とせずに、読み取り/書き込み操作ができます。

- アドインとドキュメント内のデータの間にリレーションシップが確立されます。バインドはドキュメント内に保持され、後でアクセスできます。

また、バインドを確立すると、ドキュメントまたはスプレッドシートの特定の領域を範囲とする、データおよび選択範囲の変更イベントをサブスクライブできます。つまり、ドキュメントまたはスプレッドシート全体の全般的な変更ではなく、バインドされた領域内で発生する変更のみがアドインに通知されます。

[Bindings](/javascript/api/office/office.bindings) オブジェクトが公開している [getAllAsync](/javascript/api/office/office.bindings#getallasync-options--callback-) メソッドを使用すると、ドキュメントまたはスプレッドシートで確立されている一連のすべてのバインドにアクセスできます。 個々のバインドに ID でアクセスするには、[Bindings.getBindingByIdAsync](/javascript/api/office/office.bindings#getbyidasync-id--options--callback-) メソッドまたは [Office.select](/javascript/api/office) メソッドを使用します。 新しいバインドを確立したり、既存のバインドを削除したりするには、 `Bindings` [Addfromselectionasync](/javascript/api/office/office.bindings#addfromselectionasync-bindingtype--options--callback-)、 [addFromPromptAsync](/javascript/api/office/office.bindings#addfrompromptasync-bindingtype--options--callback-)、 [addfromnameditemasync](/javascript/api/office/office.bindings#addfromnameditemasync-itemname--bindingtype--options--callback-)、または [releasebyidasync](/javascript/api/office/office.bindings#releasebyidasync-id--options--callback-)のいずれかのメソッドを使用します。

またはメソッドを使用してバインドを作成するときに、  _Bindingtype_ パラメーターで指定するバインドには3つの種類があり `addFromSelectionAsync` `addFromPromptAsync` `addFromNamedItemAsync` ます。



|**バインドの種類**|**説明**|**ホスト アプリケーションのサポート**|
|:-----|:-----|:-----|
|テキスト バインド|テキストとして表現できるドキュメントの領域にバインドします。|Word では、連続する選択範囲の大部分が有効ですが、Excel では、単一セルの範囲のみがテキスト バインドの対象です。Excel では、プレーン テキストのみがサポートされます。Word では、3 つの形式 (プレーン テキスト、HTML、および Open XML for Office) がサポートされます。|
|マトリックス バインド|ヘッダーがない表形式のデータが含まれるドキュメントの固定領域にバインドします。マトリックス バインド内のデータは、2 次元の **Array** として書き込みまたは読み取りが行われます。JavaScript では、これは、配列の配列として実装されています。たとえば、2 列の **string** 値が 2 行ある場合は ` [['a', 'b'], ['c', 'd']]` のように書き込みまたは読み取りが行われ、1 列が 3 行ある場合は `[['a'], ['b'], ['c']]` のように書き込みまたは読み取りが行われます。|Excel では、セルの連続する選択範囲を使用してマトリックス バインドを確立できます。Word では、表のみがマトリックス バインドをサポートします。|
|テーブル バインド|ヘッダーがある表が含まれるドキュメントの領域にバインドします。 テーブル バインド内のデータは、[TableData](/javascript/api/office/office.tabledata) オブジェクトとして書き込みまたは読み取りが行われます。 オブジェクトは、 `TableData` **headers** および **rows** プロパティを使用してデータを公開します。|Excel または Word の表はすべて、テーブル バインドの基礎にできます。テーブル バインドを確立すると、ユーザーが表に追加する新しい各行または各列が、自動的にバインドに含まれます。 |

<br/>

オブジェクトの3つの "add" メソッドのいずれかを使用してバインドを作成したら、 `Bindings` 対応するオブジェクトのメソッドを使用してバインドのデータとプロパティを操作できます。 [MatrixBinding](/javascript/api/office/office.matrixbinding)、 [Tablebinding](/javascript/api/office/office.tablebinding)、または [textbinding](/javascript/api/office/office.textbinding)。 これらの3つのオブジェクトはすべて、バインドされたデータを操作できるようにするオブジェクトの [getdataasync](/javascript/api/office/office.binding#getdataasync-options--callback-) メソッドと [setdataasync](/javascript/api/office/office.binding#setdataasync-data--options--callback-) メソッドを継承し `Binding` ます。

バインドに関する操作の実行方法を示すコード例については、「[ドキュメントまたはスプレッドシート内の領域へのバインド](bind-to-regions-in-a-document-or-spreadsheet.md)」を参照してください。


## <a name="working-with-custom-xml-parts-using-the-customxmlparts-and-customxmlpart-objects"></a>CustomXmlParts オブジェクトおよび CustomXmlPart オブジェクトによるカスタム XML パーツの操作


 **適用対象:** Word の作業ウィンドウ アドイン

API の [CustomXmlParts](/javascript/api/office/office.customxmlparts) オブジェクトと [CustomXmlPart](/javascript/api/office/office.customxmlpart) オブジェクトを使用すると、Word 文書内のカスタム XML パーツにアクセスできます。これにより、文書のコンテンツに対する XML 主導の操作が可能になります。 およびオブジェクトを操作するデモについ `CustomXmlParts` `CustomXmlPart` ては、「 [カスタム XML](https://github.com/OfficeDev/Word-Add-in-Work-with-custom-XML-parts) を使用して作業する」のコードサンプルを参照してください。


## <a name="working-with-the-entire-document-using-the-getfileasync-method"></a>getFileAsync メソッドを使用したドキュメント全体の操作


 **適用対象:** Word および PowerPoint の作業ウィンドウ アドイン

[Document.getFileAsync](/javascript/api/office/office.document#getfileasync-filetype--options--callback-) メソッド、および [File](/javascript/api/office/office.file) オブジェクトと [Slice](/javascript/api/office/office.slice) オブジェクトのメンバーは、一度に最大で 4 MB ずつのスライス (チャンク) に分割して Word および PowerPoint ドキュメント ファイル全体を取得する機能を提供します。詳細については、「[PowerPoint または Word 用アドインからドキュメント全体を取得する](../word/get-the-whole-document-from-an-add-in-for-word.md)」を参照してください。


## <a name="mailbox-object"></a>Mailbox オブジェクト

**適用対象:** Outlook アドイン

Outlook アドインでは、主に [Mailbox](/javascript/api/outlook/office.mailbox) オブジェクトにより公開されている API のサブセットを使用します。Outlook アドイン専用のオブジェクトおよびメンバー (たとえば、[Item](/javascript/api/outlook/office.item) オブジェクトなど) にアクセスするには、次のコード行に示すように、[Context](/javascript/api/office/office.context#mailbox) オブジェクトの **mailbox** プロパティを使用して、**Mailbox** オブジェクトにアクセスします。

```js
// Access the Item object.
var item = Office.context.mailbox.item;

```

さらに、Outlook アドインでは次のオブジェクトを使用できます。

- `Office` オブジェクト: 初期化用。

- `Context` オブジェクト: コンテンツおよび表示言語のプロパティへのアクセスに使用します。

- `RoamingSettings` オブジェクト: Outlook アドイン固有のカスタム設定を、アドインがインストールされているユーザーのメールボックスに保存するために使用します。

Outlook アドインでの JavaScript の使用については、「[Outlook アドイン](../outlook/outlook-add-ins-overview.md)」を参照してください。
