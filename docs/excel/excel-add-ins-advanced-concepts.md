---
title: Excel JavaScript API を使用した高度なプログラミングの概念
description: ''
ms.date: 07/17/2019
localization_priority: Priority
ms.openlocfilehash: a4639070ed74f9beb757de7c30d1d7e32a3e63fa
ms.sourcegitcommit: da8e6148f4bd9884ab9702db3033273a383d15f0
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/20/2019
ms.locfileid: "36477755"
---
# <a name="advanced-programming-concepts-with-the-excel-javascript-api"></a>Excel JavaScript API を使用した高度なプログラミングの概念

この記事では、「[Excel JavaScript API を使用した基本的なプログラミングの概念](excel-add-ins-core-concepts.md)」の情報を基にして、より高度な概念をいくつか説明します。これらは Excel 2016 以降の複雑なアドインを構築するために不可欠です。

## <a name="officejs-apis-for-excel"></a>Excel 用の Office.js API

Excel アドインは、次の 2 つの JavaScript オブジェクト モデルを含む JavaScript API for Office を使用して、Excel のオブジェクトを操作します。

* **Excel JavaScript API**:Office 2016 で導入された [Excel JavaScript API](/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview) には、ワークシート、範囲、表、グラフなどへのアクセスに使用できる、厳密に型指定されたオブジェクトが用意されています。

* **共通 API**: Office 2013 で導入された[共通 API](/javascript/api/office) を使用すると、複数の種類の Office アプリケーション間で共通の UI、ダイアログ、クライアント設定などの機能にアクセスすることができます。

Excel 2016 以降を対象にしたアドインでは、機能の大部分を Excel JavaScript API を使用して開発する可能性がありますが、共通 API のオブジェクトも使用します。 次に例を示します。

- [Context](/javascript/api/office/office.context): **Context** オブジェクトは、アドインのランタイム環境を表し、API の主要なオブジェクトへのアクセスを提供します。 これは `contentLanguage` や `officeTheme` などのブック構成の詳細で構成され、`host` や `platform` などのアドインのランタイム環境に関する情報も提供します。 さらに、`requirements.isSetSupported()` メソッドも提供されます。これを使用すると、指定した要件セットが、アドインが実行されている Excel アプリケーションでサポートされているかどうかを確認できます。

- [Document](/javascript/api/office/office.document):**Document** オブジェクトは `getFileAsync()` メソッドを提供します。これを使用すると、アドインが実行されている Excel ファイルをダウンロードできます。

## <a name="requirement-sets"></a>要件セット

要件セットは、API メンバーの名前付きグループです。 Office アドインはランタイム チェックを実行できます。または、マニフェストで指定されている要件セットを使用して、Office ホストがアドインに必要な API をサポートしているかどうかを確認できます。 サポートされている各プラットフォームで使用できる特定の要件セットを確認するには、「[Excel JavaScript API の要件セット](/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets)」を参照してください。

### <a name="checking-for-requirement-set-support-at-runtime"></a>実行時に要件セットのサポートを確認する

次のコード サンプルは、アドインが実行されているホスト アプリケーションが指定された API の要件セットをサポートしているかどうかを確認する方法を示しています。

```js
if (Office.context.requirements.isSetSupported('ExcelApi', '1.3')) {
  /// perform actions
}
else {
  /// provide alternate flow/logic
}
```

### <a name="defining-requirement-set-support-in-the-manifest"></a>マニフェストで要件セットのサポートを定義する

アドインのマニフェストで [Requirements 要素](/office/dev/add-ins/reference/manifest/requirements) を使用して、アドインをアクティブにするために必要な最小要件セットや API メソッド (またはその両方) を指定できます。 Office ホストまたはプラットフォームが、マニフェストの **Requirements** 要素で指定した要件セットまたは API メソッドをサポートしない場合、アドインはそのホストまたはプラットフォームでは実行されず、**[個人用アドイン]** に表示されるアドインの一覧にも表示されません。

次のコード サンプルは、アドインが ExcelApi 要件セットのバージョン 1.3 以上をサポートする Office ホスト アプリケーションのすべて読み込まれる必要があることを指定する、アドインのマニフェストの **Requirements** 要素を示しています。

```xml
<Requirements>
   <Sets DefaultMinVersion="1.3">
      <Set Name="ExcelApi" MinVersion="1.3"/>
   </Sets>
</Requirements>
```

> [!NOTE]
> Excel on the web、Windows、iPad などの Office ホストのプラットフォームすべてでアドインを使用できるようにするには、マニフェストで要件セットのサポートを定義するのではなく、実行時に要件のサポートを確認することをお勧めします。

### <a name="requirement-sets-for-the-officejs-common-api"></a>Office.js 共通 API の要件セット

共通 API の要件セットの詳細については、「[Office 共通 API の要件セット](/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)」をご覧ください。

## <a name="loading-the-properties-of-an-object"></a>オブジェクトのプロパティを読み込む

Excel JavaScript オブジェクトで `load()` メソッドを呼び出すと、API は `sync()` メソッドの実行時にオブジェクトを JavaScript メモリに読み込むように指示されます。 `load()` メソッドには、読み込むプロパティのコンマで区切られた名前を含む文字列や、読み込むプロパティを指定するオブジェクト、改ページのオプションなどを指定できます。

> [!NOTE]
> パラメーターを指定せずにオブジェクト (またはコレクション) の `load()` メソッドを呼び出すと、オブジェクトのすべてのスカラー プロパティ (またはコレクション内のすべてのオブジェクトのすべてのスカラー プロパティ) が読み込まれます。 Excel ホスト アプリケーションとアドイン間のデータ転送量を減らすには、読み込むプロパティを明示的に指定しないで `load()` メソッドを呼び出さないようにします。

### <a name="method-details"></a>メソッドの詳細

#### <a name="loadparam-object"></a>load(param: object)

JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。

#### <a name="syntax"></a>構文

```js
object.load(param);
```

#### <a name="parameters"></a>パラメーター

|**パラメーター**|**型**|**説明**|
|:------------|:-------|:----------|
|`param`|object|省略可能。プロパティ名を、コンマで区切られた文字列または 1 つの配列として指定します。 オブジェクトを渡して、選択プロパティとナビゲーション プロパティを設定することもできます (次の例を参照)。|

#### <a name="returns"></a>戻り値

void

#### <a name="example"></a>例

次のコード サンプルでは、別の範囲のプロパティをコピーして 1 つの Excel 範囲のプロパティを設定します。 プロパティ値にアクセスして対象範囲に書き込む前に、ソース オブジェクトを最初に読み込む必要があることに注意してください。 この例では、2 つの範囲 (**B2:E2** および **B7:E7**) のデータがあり、2 つの範囲の書式設定が最初は異なっていると仮定します。

```js
Excel.run(function (ctx) {
    var sheet = ctx.workbook.worksheets.getItem("Sample");
    var sourceRange = sheet.getRange("B2:E2");
    sourceRange.load("format/fill/color, format/font/name, format/font/color");

    return ctx.sync()
        .then(function () {
            var targetRange = sheet.getRange("B7:E7");
            targetRange.set(sourceRange);
            targetRange.format.autofitColumns();

            return ctx.sync();
        });
}).catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

### <a name="load-option-properties"></a>オプションのプロパティを読み込む

`load()` メソッドを呼び出すときに、コンマで区切られた文字列または配列を渡す代わりに、次のプロパティを含むオブジェクトを渡すことができます。

|**プロパティ**|**型**|**説明**|
|:-----------|:-------|:----------|
|`select`|object|スカラー プロパティ名のコンマ区切りリストまたは配列が含まれています。省略可能。|
|`expand`|object|ナビゲーション プロパティ名のコンマ区切りリストまたは配列が含まれています。省略可能。|
|`top`|int| 結果に含めることができるコレクション項目の最大数を指定します。省略可能。このオプションは、オブジェクト表記オプションを使用する場合にのみ使用できます。|
|`skip`|int|スキップされて結果に組み込まれないコレクション内の項目の数を指定します。`top` が指定されている場合は、指定された数の項目がスキップされた後で結果セットが開始されます。省略可能。このオプションは、オブジェクト表記オプションを使用する場合にのみ使用できます。|

次のコードサンプルは、`name` プロパティと `address`コレクション内の各ワークシートの使用範囲を選択して、ワークシートコレクションを読み込みます。 また、コレクションの上位 5 つのワークシートのみを読み込むように指定しています。 `top: 10` と `skip: 5` を属性値として指定することで、次の 5 つのワークシートのセットを処理できます。

```js
myWorksheets.load({
    select: 'name, userRange/address',
    expand: 'tables',
    top: 5,
    skip: 0
});
```

## <a name="scalar-and-navigation-properties"></a>スカラー プロパティとナビゲーション プロパティ

プロパティには、**スカラー**と**ナビゲーション**という 2 つのカテゴリがあります。 スカラー プロパティは、文字列、整数、JSON 構造体などの割り当て可能な型です。 ナビゲーション プロパティは、プロパティを直接割り当てるのではなく、読み取り専用のオブジェクトと、そのフィールドが割り当てられているオブジェクトのコレクションです。 たとえば、[ワークシート](/javascript/api/excel/excel.worksheet) オブジェクトの `name` メンバーと `position` メンバーはスカラー プロパティですが、`protection` と `tables` はナビゲーション プロパティです。 [DataValidation] オブジェクトの `prompt` は、サブプロパティ (`dv.prompt.title = "MyPrompt" // will not set the title`) を設定するのではなく、JSON オブジェクト (`dv.prompt = { title: "MyPrompt"}`) を使用して設定する必要があるスカラー プロパティの例です。

### <a name="scalar-properties-and-navigation-properties-with-objectload"></a>`object.load()` を使用したスカラー プロパティとナビゲーション プロパティ

パラメーターを指定しないで `object.load()` メソッドを呼び出すと、オブジェクトのすべてのスカラー プロパティが読み込まれます。オブジェクトのナビゲーション プロパティは読み込まれません。 さらに、ナビゲーション プロパティを直接読み込むことはできません。 その代わりに、`load()` メソッドを使用して、目的のナビゲーション プロパティ内のスカラー プロパティを個別に参照する必要があります。 たとえば、範囲のフォント名を読み込むには、**name** プロパティへのパスとして **format** および **font** ナビゲーション プロパティを指定する必要があります。

```js
someRange.load("format/font/name")
```

> [!NOTE]
> Excel JavaScript API を使用すると、パスを詳しく調べることでナビゲーション プロパティのスカラー プロパティを設定できます。 たとえば、`someRange.format.font.size = 10;` を使用して範囲のフォント サイズを設定できます。 設定前にプロパティを読み込む必要はありません。 

## <a name="setting-properties-of-an-object"></a>オブジェクトのプロパティを設定する

入れ子になったナビゲーション プロパティを持つオブジェクトのプロパティを設定するのは面倒です。 前述のナビゲーション パスを使用してプロパティを個別に設定する代わりに、Excel JavaScript API のすべてのオブジェクトで使用できる、`object.set()` メソッドを使用できます。 このメソッドを使用すると、同じ Office.js 型の別のオブジェクト、またはメソッドが呼び出されるオブジェクトのプロパティと同様に構造化されたプロパティを持つ JavaScript オブジェクトを渡すことによって、オブジェクトの複数のプロパティを一度に設定できます。

> [!NOTE]
> `set()` メソッドは、Excel JavaScript API などホスト固有の Office JavaScript API のオブジェクトでのみ実装されます。 共通 (共有) API は、このメソッドをサポートしていません。 

### <a name="set-properties-object-options-object"></a>set (properties: object, options: object)

メソッドが呼び出されるオブジェクトのプロパティは、渡されたオブジェクトの対応するプロパテに指定された値に設定されます。`properties` パラメーターが JavaScript オブジェクトの場合、メソッドが呼び出される読み取り専用プロパティに対応する渡されたオブジェクトの任意のプロパティは、`options` パラメーターの値に応じて、無視されるか、例外のスローが発生します。

#### <a name="syntax"></a>構文

```js
object.set(properties[, options]);
```

#### <a name="parameters"></a>パラメーター

|**パラメーター**|**型**|**説明**|
|:------------|:--------|:----------|
|`properties`|object|メソッドが呼び出されるオブジェクトの同じ Office.js 型のオブジェクト、またはメソッドが呼び出されるオブジェクトの構造を反映するプロパティ名と型を持つ JavaScript オブジェクトのいずれかです。|
|`options`|object|省略可能。最初のパラメーターが JavaScript オブジェクトの場合にのみ渡すことができます。オブジェクトには、次のプロパティを含めることができます。`throwOnReadOnly?: boolean` (既定値は `true`。渡された JavaScript オブジェクトに読み取り専用プロパティが含まれている場合は、エラーをスローします。)|

#### <a name="returns"></a>戻り値

void

#### <a name="example"></a>例

次のコード サンプルは、`set()` メソッドを呼び出し、**Range** オブジェクトのプロパティの構造を反映するプロパティ名と型を持つ JavaScript オブジェクトを渡すことによって、範囲のいくつかの書式プロパティを設定します。この例では、範囲 **B2:E2** にデータがあると仮定します。

```js
Excel.run(function (ctx) {
    var sheet = ctx.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B2:E2");
    range.set({
        format: {
            fill: {
                color: '#4472C4'
            },
            font: {
                name: 'Verdana',
                color: 'white'
            }
        }
    });
    range.format.autofitColumns();

    return ctx.sync();
}).catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

## <a name="42ornullobject-methods"></a>&#42;OrNullObject メソッド

多くの Excel JavaScript API メソッドは、API の条件が満たされない場合に例外を返します。 たとえば、ブックに存在しないワークシート名を指定してワークシートを取得しようとすると、`getItem()` メソッドは `ItemNotFound` 例外を返します。 

このようなシナリオの複雑な例外処理ロジックを実装する代わりに、Excel JavaScript API のいくつかのメソッドで使用できる `*OrNullObject` メソッドのバリエーションを使用できます。 指定された項目が存在しない場合、`*OrNullObject` メソッドは例外をスローするのではなく、null オブジェクト (JavaScript `null` ではない) を返します。 たとえば、`getItemOrNullObject()` などのコレクションで **** メソッドを呼び出して、コレクションからのアイテムの取得を試行できます。 `getItemOrNullObject()` メソッドは、指定された項目が存在する場合はその項目を返し、それ以外の場合は null オブジェクトを返します。 返される null オブジェクトには、ブール型プロパティ `isNullObject` が含まれています。これを評価して、オブジェクトが存在するかどうかを判断できます。

次のコード サンプルは `getItemOrNullObject()` メソッドを使用して、"Data" という名前のワークシートの取得を試行します。 メソッドが null オブジェクトを返す場合は、新しいシートを作成し、そのシート上で操作を実行する必要があります。

```js
var dataSheet = context.workbook.worksheets.getItemOrNullObject("Data");

return context.sync()
  .then(function() {
    if (dataSheet.isNullObject) {
        // Create the sheet
    }

    dataSheet.position = 1;
    //...
  })
```

## <a name="see-also"></a>関連項目

* [Excel JavaScript API を使用した基本的なプログラミングの概念](excel-add-ins-core-concepts.md)
* [Excel アドインのコード サンプル](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)
* [Excel の JavaScript API を使用した、パフォーマンスの最適化](performance.md)
* [Excel JavaScript API リファレンス](/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview)
