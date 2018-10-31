---
title: Excel JavaScript API を使用した高度なプログラミングの概念
description: ''
ms.date: 10/03/2018
ms.openlocfilehash: 09f2d95e4cf7631b519f00cddee265dbf697e07e
ms.sourcegitcommit: c53f05bbd4abdfe1ee2e42fdd4f82b318b363ad7
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/12/2018
ms.locfileid: "25505889"
---
# <a name="advanced-programming-concepts-with-the-excel-javascript-api"></a>Excel JavaScript API を使用した高度なプログラミングの概念

この記事では、「[ Excel JavaScript API の基本的なプログラミングの概念](excel-add-ins-core-concepts.md) 」の情報を基にして、より高度な概念をいくつか説明します。これらは Excel 2016 の複雑なアドインを構築するために不可欠です。

## <a name="officejs-apis-for-excel"></a>Excel 用の Office.js API

Excel アドインは、次の 2 つの JavaScript オブジェクト モデルを含む JavaScript API for Office を使用して、Excel のオブジェクトを操作します。

* **Excel JavaScript API**: Office 2016 で導入された [Excel JavaScript API](https://docs.microsoft.com/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview?view=office-js) には、ワークシート、範囲、表、グラフなどへのアクセスに使用できる、厳密に型指定されたオブジェクトが用意されています。 

* **共通 API**: Office 2013 で導入された共通 API ([共有 API](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office?view=office-js) とも呼ばれる) を使用すると、Word、Excel、PowerPoint など複数の種類のホスト アプリケーションに共通する UI、ダイアログ、クライアント設定などの機能にアクセスできます。

Excel 2016 を対象にしたアドインでは、機能の大部分を Excel JavaScript API を使用して開発する可能性がありますが、共有 API のオブジェクトも使用します。例:

- [コンテキスト](https://docs.microsoft.com/javascript/api/office/office.context?view=office-js): **コンテキスト** オブジェクトは、アドインのランタイム環境を表し、API の主要なオブジェクトへのアクセスを提供します。`contentLanguage` や `officeTheme` のようなブック構成の詳細を含み、`host` と `platform` のようなアドインのランタイム環境に関する情報も提供します。さらに、`requirements.isSetSupported()` メソッドを提供し、指定された要件セットがアドインが実行されている Excel のアプリケーションでサポートされているかを確認するために使用することができます。 

- [Document](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js):**Document** オブジェクトは `getFileAsync()` メソッドを提供します。これを使用すると、アドインが実行されている Excel ファイルをダウンロードできます。 

## <a name="requirement-sets"></a>要件セット

要件のセットを API メンバーのグループと呼びます。Office アドインを実行時チェックを実行したり、Office ホストがアドインを必要とする Api をサポートしているかどうかを判断するには、マニフェストで指定されている要件のセットを使用できます。サポートされる各プラットフォームで利用可能な特定の要件のセットを識別するには、 [Excel の JavaScript API の要件の設定](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets?view=office-js)を参照してください。

### <a name="checking-for-requirement-set-support-at-runtime"></a>実行時に要件セットのサポートを確認する

次のコード サンプルは、アドインが実行されているホスト アプリケーションが指定された API の要件セットをサポートしているかどうかを確認する方法を示しています。

```js
if (Office.context.requirements.isSetSupported('ExcelApi', 1.3) === true) {
  /// perform actions
}
else {
  /// provide alternate flow/logic
}
```

### <a name="defining-requirement-set-support-in-the-manifest"></a>マニフェストで要件セットのサポートを定義する

アドインのマニフェストで [要件の要素](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/requirements?view=office-js) を使用して、最小限の要件セットおよび/またはアドインを有効にするのに必要な API メソッドを指定します。Office ホストまたはプラットフォームが要件セットまたはマニフェストの **要件** の要素で指定されている API のメソッドをサポートしていない場合は、アドインはそのホストまたはプラットフォームでは実行されず、 **My アドイン** に表示されるアドインの一覧に表示されません。 

次のコード サンプルは、アドインが ExcelApi 要件セットのバージョン 1.3 以上をサポートする Office ホスト アプリケーションのすべて読み込まれる必要があることを指定する、アドインのマニフェストの **Requirements** 要素を示しています。

```xml
<Requirements>
   <Sets DefaultMinVersion="1.3">
      <Set Name="ExcelApi" MinVersion="1.3"/>
   </Sets>
</Requirements>
```

> [!NOTE]
> Excel for Windows、Excel Online、Excel for iPad などの Office ホストのプラットフォームすべてでアドインを使用できるようにするには、マニフェストで要件セットのサポートを定義するのではなく、実行時に要件のサポートを確認することをお勧めします。

### <a name="requirement-sets-for-the-officejs-common-api"></a>Office.js 共通 API の要件セット

共通 API の要件セットの詳細は、「[Office 共通 API の要件セット](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets?view=office-js)」を参照してください。

## <a name="loading-the-properties-of-an-object"></a>オブジェクトのプロパティを読み込む

Excel JavaScript オブジェクトで `load()` メソッドを呼び出すと、API は`sync()` メソッドの実行時にオブジェクトを JavaScript メモリに読み込むように指示されます。`load()` メソッドには、読み込むプロパティのコンマで区切られた名前を含む文字列や、読み込むプロパティを指定するオブジェクト、改ページのオプションなどを指定できます。 

> [!NOTE]
> パラメーターを指定せずにオブジェクト (またはコレクション) の `load()` メソッドを呼び出すと、オブジェクトのすべてのスカラー プロパティ (またはコレクション内のすべてのオブジェクトのすべてのスカラー プロパティ) が読み込まれます。Excel ホスト アプリケーションとアドイン間のデータ転送量を減らすには、読み込むプロパティを明示的に指定しないで `load()`  メソッドを呼び出さないようにします。

### <a name="method-details"></a>メソッドの詳細

#### <a name="loadparam-object"></a>load(param: object)

JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。

#### <a name="syntax"></a>構文

```js
object.load(param);
```

#### <a name="parameters"></a>パラメーター

|**パラメーター**|**種類**|**説明**|
|:------------|:-------|:----------|
|`param`|object|省略可能です。パラメーターとの関係の名前をコンマで区切られた文字列または配列を受け取ります。(次の例で示すように) 選択とナビゲーション プロパティを設定するオブジェクトを渡すこともできます。|

#### <a name="returns"></a>戻り値

Void

#### <a name="example"></a>例

次のコード サンプルでは、別の範囲のプロパティをコピーして 1 つの Excel の範囲のプロパティを設定します。プロパティ値にアクセスして対象範囲に書き込む前に、ソース オブジェクトを最初に読み込む必要があることに注意してください。この例では、2 つの範囲 (**B2:E2** および **B7:E7**) のデータがあり、2 つの範囲の書式設定が最初は異なっていると仮定します。

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

|**プロパティ**|**種類**|**説明**|
|:-----------|:-------|:----------|
|`select`|object|パラメーター/リレーションシップの名前のコンマ区切りリストまたは配列が含まれます。省略可能。|
|`expand`|object|リレーションシップ名のコンマ区切りリストまたは配列が含まれています。省略可能。|
|`top`|int| 結果に含めることができるコレクション項目の最大数を指定します。省略可能。このオプションは、オブジェクト表記オプションを使用する場合にのみ使用できます。|
|`skip`|int|スキップされて結果に組み込まれないコレクション内の項目の数を指定します。`top` が指定されている場合は、指定された数の項目がスキップされた後で結果セットが開始されます。省略可能。このオプションは、オブジェクト表記オプションを使用する場合にのみ使用できます。|

次のコード サンプルは、コレクション内の各ワークシートの使用範囲の `name` プロパティと `address` を選択することにより、ワークシートのコレクションを読み込みます。また、コレクションの最上位の 5 つのワークシートのみを読み込むことを指定します。`top: 10` と `skip: 5`   を属性値として指定することで、次の 5 つのワークシートのセットを処理できます。 

```js 
myWorksheets.load({
    select: 'name, userRange/address',
    expand: 'tables',
    top: 5,
    skip: 0
});
```

## <a name="scalar-and-navigation-properties"></a>スカラー プロパティとナビゲーション プロパティ 

Excel JavaScript API のリファレンス ドキュメントでは、オブジェクトのメンバーは、2 つのカテゴリにグループ化されています: **プロパティ** と **リレーションシップ**です。オブジェクトのプロパティは、文字列、整数、ブール値などのスカラー メンバーです。一方、オブジェクトのリレーションシップ (ナビゲーション プロパティとも呼ばれる) は、オブジェクトまたはオブジェクトのコレクションのいずれかであるメンバーです。たとえば、[  Worksheet](https://docs.microsoft.com/javascript/api/excel/excel.worksheet?view=office-js) オブジェクトの `name` メンバーと `position` メンバーはスカラー プロパティですが、`protection` と `tables` はリレーションシップ (ナビゲーション プロパティ) です。 

### <a name="scalar-properties-and-navigation-properties-with-objectload"></a>を使用したスカラー プロパティとナビゲーション プロパティ `object.load()`

パラメーターを指定しないで `object.load()` メソッドを呼び出すと、オブジェクトのすべてのスカラー プロパティが読み込まれます。オブジェクトのナビゲーション プロパティは読み込まれません。さらに、ナビゲーション プロパティは直接読み込むことができません。代わりに、`load()` メソッドを使用して、目的のナビゲーション プロパティ内の個別のスカラー プロパティを参照する必要があります。たとえば、範囲のフォント名を読み込むには、**name** プロパティへのパスとして **format** および **font** ナビゲーション プロパティを指定する必要があります。

```js
someRange.load("format/font/name")
```

> [!NOTE]
> Excel JavaScript API を使用すると、パスを詳しく調べることでナビゲーション プロパティのスカラー プロパティを設定できます。たとえば、`someRange.format.font.size = 10;` を使用して範囲のフォント サイズを設定できます。プロパティを設定する前にロードする必要はありません。 

## <a name="setting-properties-of-an-object"></a>オブジェクトのプロパティを設定する

入れ子になったナビゲーション プロパティを持つオブジェクトのプロパティの設定は面倒です。前述のナビゲーション パスを使用してプロパティを個別に設定する代わりに、Excel JavaScript API のすべてのオブジェクトで使用できる、`object.set()` メソッドを使用できます。このメソッドを使用すると、同じ Office.js 型の別のオブジェクト、またはメソッドが呼び出されるオブジェクトのプロパティと同様に構造化されたプロパティを持つ JavaScript オブジェクトを渡すことによって、オブジェクトの複数のプロパティを一度に設定できます。

> [!NOTE]
> `set()`メソッドは、Excel JavaScript API などホスト固有の Office JavaScript API のオブジェクトでのみ実装されます。共通 (共有) API は、このメソッドをサポートしていません。 

### <a name="set-properties-object-options-object"></a>set (properties: object, options: object)

メソッドが呼び出されるオブジェクトのプロパティは、渡されたオブジェクトの対応するプロパテに指定された値に設定されます。`properties` パラメーターが JavaScript オブジェクトの場合、メソッドが呼び出される読み取り専用プロパティに対応する渡されたオブジェクトの任意のプロパティは、`options` パラメーターの値に応じて、無視されるか、例外のスローが発生します。

#### <a name="syntax"></a>構文

```js
object.set(properties[, options]);
```

#### <a name="parameters"></a>パラメーター

|**パラメーター**|**種類**|**説明**|
|:------------|:--------|:----------|
|`properties`|object|メソッドが呼び出されるオブジェクトの同じ Office.js 型のオブジェクト、またはメソッドが呼び出されるオブジェクトの構造を反映するプロパティ名と型を持つ JavaScript オブジェクトのいずれかです。|
|`options`|object|省略可能。最初のパラメーターが JavaScript オブジェクトの場合にのみ渡すことができます。オブジェクトには、次のプロパティを含めることができます。`throwOnReadOnly?: boolean` (既定値は `true`。渡された JavaScript オブジェクトに読み取り専用プロパティが含まれている場合は、エラーをスローします。)|

#### <a name="returns"></a>戻り値

Void    

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
## <a name="42ornullobject-methods"></a>*OrNullObject メソッド

多くの Excel JavaScript API メソッドは、API の条件が満たされない場合に例外を返します。たとえば、ブックに存在しないワークシート名を指定してワークシートを取得しようとすると、`getItem()` メソッドは `ItemNotFound` 例外を返します。 

このようなシナリオの複雑な例外処理論理を実装する代わりに、Excel JavaScript API のいくつかのメソッドで使用可能な`*OrNullObject`メソッド変数を使用することができます。`*OrNullObject`メソッドは、特定のアイテムが存在しなければ例外を投げずに null 値 (JavaScript`null` ではない) を返します。たとえば、**ワークシート** のようなコレクションで`getItemOrNullObject()`メソッドを呼び出して、コレクションからアイテムを引き出すことができます。`getItemOrNullObject()`メソッドは指定したアイテムが存在すればそれを返しますが、そうでなければ null オブジェクトを返します。返される null オブジェクトには、オブジェクトが存在するかの決定を評価できる boolean プロパティ`isNullObject`が含まれます。

次のコード サンプルは `getItemOrNullObject()` メソッドを使用して、"Data" という名前のワークシートの取得を試行します。メソッドが null オブジェクトを返す場合は、新しいシートを作成し、そのシート上で操作を実行する必要があります。

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
* [Excel JavaScript API パフォーマンスの最適化](performance.md)
* [Excel JavaScript API リファレンス](https://docs.microsoft.com/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview?view=office-js)
