---
title: JavaScript API を使用して図形Excelする
description: 図形をExcel図面レイヤーに配置するオブジェクトとして定義する方法についてExcel。
ms.date: 01/14/2020
localization_priority: Normal
ms.openlocfilehash: 533a9cf9689bcaa5cd43635da836730a2af6ab61
ms.sourcegitcommit: 3fa8c754a47bab909e559ae3e5d4237ba27fdbe4
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/30/2021
ms.locfileid: "53671472"
---
# <a name="work-with-shapes-using-the-excel-javascript-api"></a>JavaScript API を使用して図形Excelする

Excel図形は、図形の描画レイヤーに配置されるオブジェクトとして定義Excel。 つまり、セルの外側にあるものは図形です。 この記事では、図形および[ShapeCollection](/javascript/api/excel/excel.shapecollection) API と組み合[](/javascript/api/excel/excel.shape)わせて幾何学的図形、線、および画像を使用する方法について説明します。 [グラフ](/javascript/api/excel/excel.chart)については、独自の記事[「JavaScript API](excel-add-ins-charts.md)を使用してグラフをExcelします。

次の図は、体温計を形成する図形を示しています。
![図形として作成された体温計Excel。](../images/excel-shapes.png)

## <a name="create-shapes"></a>図形を作成する

図形は、ワークシートの図形コレクション () を通じて作成され、格納されます `Worksheet.shapes` 。 `ShapeCollection` この目的 `.add*` のためにいくつかのメソッドがあります。 すべての図形には、コレクションに追加するときに名前と ID が生成されます。 これらは、それぞれ `name` プロパティ `id` とプロパティです。 `name` メソッドを使用して簡単に取得できるようアドインで設定 `ShapeCollection.getItem(name)` できます。

次の種類の図形は、関連付けられたメソッドを使用して追加されます。

| Shape | Tabs.Add メソッド (Outlook フォーム スクリプト) | 署名 |
|-------|------------|-----------|
| ジオメトリック シェイプ | [addGeometricShape](/javascript/api/excel/excel.shapecollection#addGeometricShape_geometricShapeType_) | `addGeometricShape(geometricShapeType: Excel.GeometricShapeType): Excel.Shape` |
| 画像 (JPEG または PNG) | [addImage](/javascript/api/excel/excel.shapecollection#addImage_base64ImageString_) | `addImage(base64ImageString: string): Excel.Shape` |
| Line | [addLine](/javascript/api/excel/excel.shapecollection#addLine_startLeft__startTop__endLeft__endTop__connectorType_) | `addLine(startLeft: number, startTop: number, endLeft: number, endTop: number, connectorType?: Excel.ConnectorType): Excel.Shape` |
| SVG | [addSvg](/javascript/api/excel/excel.shapecollection#addSvg_xml_) | `addSvg(xml: string): Excel.Shape` |
| テキスト ボックス | [addTextBox](/javascript/api/excel/excel.shapecollection#addTextBox_text_) | `addTextBox(text?: string): Excel.Shape` |

### <a name="geometric-shapes"></a>幾何学的図形

で幾何学的な図形が作成されます `ShapeCollection.addGeometricShape` 。 このメソッドは、 [引数として GeometricShapeType](/javascript/api/excel/excel.geometricshapetype) 列挙型を受け取ります。

次のコード サンプルでは、ワークシートの上辺と左側から 100 ピクセルの位置にある **"Square"** という名前の 150x150 ピクセルの四角形を作成します。

```js
// This sample creates a rectangle positioned 100 pixels from the top and left sides
// of the worksheet and is 150x150 pixels.
Excel.run(function (context) {
    var shapes = context.workbook.worksheets.getItem("MyWorksheet").shapes;
    var rectangle = shapes.addGeometricShape(Excel.GeometricShapeType.rectangle);
    rectangle.left = 100;
    rectangle.top = 100;
    rectangle.height = 150;
    rectangle.width = 150;
    rectangle.name = "Square";
    return context.sync();
}).catch(errorHandlerFunction);
```

### <a name="images"></a>画像

JPEG、PNG、SVG 画像は、ワークシートに図形として挿入できます。 メソッド `ShapeCollection.addImage` は引数として base64 でエンコードされた文字列を受け取ります。 これは、文字列形式の JPEG または PNG イメージのいずれかです。 `ShapeCollection.addSvg` この引数はグラフィックを定義する XML ですが、文字列も取り込まれます。

次のコード サンプルは [、FileReader](https://developer.mozilla.org/docs/Web/API/FileReader) が文字列として読み込むイメージ ファイルを示しています。 この文字列には、図形が作成される前にメタデータ "base64" が削除されます。

```js
// This sample creates an image as a Shape object in the worksheet.
var myFile = document.getElementById("selectedFile");
var reader = new FileReader();

reader.onload = (event) => {
    Excel.run(function (context) {
        var startIndex = reader.result.toString().indexOf("base64,");
        var myBase64 = reader.result.toString().substr(startIndex + 7);
        var sheet = context.workbook.worksheets.getItem("MyWorksheet");
        var image = sheet.shapes.addImage(myBase64);
        image.name = "Image";
        return context.sync();
    }).catch(errorHandlerFunction);
};

// Read in the image file as a data URL.
reader.readAsDataURL(myFile.files[0]);
```

### <a name="lines"></a>Lines

行がで作成されます `ShapeCollection.addLine` 。 このメソッドには、行の開始点と終了点の左余白と上余白が必要です。 また [、ConnectorType 列挙型を使用](/javascript/api/excel/excel.connectortype) して、エンドポイント間の行のコントルト方法を指定します。 次のコード サンプルでは、ワークシートに直線を作成します。

```js
// This sample creates a straight line from [200,50] to [300,150] on the worksheet
Excel.run(function (context) {
    var shapes = context.workbook.worksheets.getItem("MyWorksheet").shapes;
    var line = shapes.addLine(200, 50, 300, 150, Excel.ConnectorType.straight);
    line.name = "StraightLine";
    return context.sync();
}).catch(errorHandlerFunction);
```

線は他の Shape オブジェクトに接続できます。 and メソッドは、指定した接続ポイントの図形に線の開始位置と終了 `connectBeginShape` `connectEndShape` 位置をアタッチします。 これらのポイントの位置は図形によって異なりますが、アドインが境界外のポイントに接続しない場合に使用 `Shape.connectionSiteCount` できます。 線は、and メソッドを使用して、接続されている図形 `disconnectBeginShape` から `disconnectEndShape` 切断されます。

次のコード サンプルでは **、"MyLine"** 行を **"LeftShape" と "RightShape"** という名前の 2 つの図形 **に接続します**。

```js
// This sample connects a line between two shapes at connection points '0' and '3'.
Excel.run(function (context) {
    var shapes = context.workbook.worksheets.getItem("MyWorksheet").shapes;
    var line = shapes.getItem("MyLine").line;
    line.connectBeginShape(shapes.getItem("LeftShape"), 0);
    line.connectEndShape(shapes.getItem("RightShape"), 3);
    return context.sync();
}).catch(errorHandlerFunction);
```

## <a name="move-and-resize-shapes"></a>図形の移動とサイズ変更

図形はワークシートの上に表示されます。 配置は and プロパティによって `left` 定義 `top` されます。 これらはワークシートのそれぞれのエッジの余白として機能し、[0, 0] は左上隅になります。 これらは、and メソッドを使用して、直接設定するか、現在の位置から `incrementLeft` `incrementTop` 調整できます。 既定の位置から回転する図形の量も、この方法で確立され、プロパティは絶対量であり、メソッドは既存の回転 `rotation` `incrementRotation` を調整します。

他の図形に対する図形の深さは、プロパティによって定義 `zorderPosition` されます。 これは `setZOrder` [、ShapeZOrder](/javascript/api/excel/excel.shapezorder)を受け取るメソッドを使用して設定されます。 `setZOrder` 他の図形を基準に現在の図形の順序を調整します。

アドインには、図形の高さと幅を変更するためのいくつかのオプションがあります。 またはプロパティを `height` 設定 `width` すると、他のディメンションを変更せずに指定したディメンションが変更されます。 現在のサイズまたは元のサイズ (指定された ShapeScaleType の値に基づいて) を基準に図形のそれぞれの寸法を `scaleHeight` `scaleWidth` [調整します](/javascript/api/excel/excel.shapescaletype)。 省略可能 [な ShapeScaleFrom](/javascript/api/excel/excel.shapescalefrom) パラメーターは、図形のスケールの場所 (左上隅、中央、右下隅) を指定します。 プロパティが true の場合、スケール メソッドは他の次元も調整することで、図形の現在の縦横比 `lockAspectRatio` を維持します。 

> [!NOTE]
> プロパティへの直接の `height` 変更 `width` は、プロパティの値に関係なく、そのプロパティ `lockAspectRatio` にのみ影響します。

次のコード サンプルは、元のサイズの 1.25 倍に拡大縮小され、30 度回転された図形を示しています。

```js
// In this sample, the shape "Octagon" is rotated 30 degrees clockwise
// and scaled 25% larger, with the upper-left corner remaining in place.
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("MyWorksheet");
    var shape = sheet.shapes.getItem("Octagon");
    shape.incrementRotation(30);
    shape.lockAspectRatio = true;
    shape.scaleWidth(
        1.25,
        Excel.ShapeScaleType.currentSize,
        Excel.ShapeScaleFrom.scaleFromTopLeft);
    return context.sync();
}).catch(errorHandlerFunction);
```

## <a name="text-in-shapes"></a>図形内のテキスト

幾何学的図形にはテキストを含めできます。 図形には `textFrame` 、TextFrame 型の [プロパティがあります](/javascript/api/excel/excel.textframe)。 オブジェクト `TextFrame` は、テキスト表示オプション (余白やテキスト オーバーフローなど) を管理します。 `TextFrame.textRange` は、 [テキストコンテンツと](/javascript/api/excel/excel.textrange) フォント設定を持つ TextRange オブジェクトです。

次のコード サンプルでは、"Shape Text" というテキストを持つ "Wave" という名前の幾何学的な図形を作成します。 また、図形とテキストの色を調整し、テキストの水平方向の配置を中央に設定します。

```js
// This sample creates a light-blue wave shape and adds the purple text "Shape text" to the center.
Excel.run(function (context) {
    var shapes = context.workbook.worksheets.getItem("MyWorksheet").shapes;
    var wave = shapes.addGeometricShape(Excel.GeometricShapeType.wave);
    wave.left = 100;
    wave.top = 400;
    wave.height = 50;
    wave.width = 150;
    wave.name = "Wave";
    wave.fill.setSolidColor("lightblue");
    wave.textFrame.textRange.text = "Shape text";
    wave.textFrame.textRange.font.color = "purple";
    wave.textFrame.horizontalAlignment = Excel.ShapeTextHorizontalAlignment.center;
    return context.sync();
}).catch(errorHandlerFunction);
```

この `addTextBox` メソッドは `ShapeCollection` 、白い背景 `GeometricShape` と黒いテキスト `Rectangle` を持つ型を作成します。 これは、[挿入] タブの [テキスト Excel] ボタンによって作成される操作と **同** じです。 `addTextBox`のテキストを設定する引数 string を受け取ります `TextRange` 。

次のコード サンプルは、テキスト "Hello!" を含むテキスト ボックスの作成を示しています。

```js
// This sample creates a text box with the text "Hello!" and sizes it appropriately.
Excel.run(function (context) {
    var shapes = context.workbook.worksheets.getItem("MyWorksheet").shapes;
    var textbox = shapes.addTextBox("Hello!");
    textbox.left = 100;
    textbox.top = 100;
    textbox.height = 20;
    textbox.width = 45;
    textbox.name = "Textbox";
    return context.sync();
}).catch(errorHandlerFunction);
```

## <a name="shape-groups"></a>図形グループ

図形はグループ化できます。 これにより、ユーザーはそれらを配置、サイズ変更、その他の関連タスク用の 1 つのエンティティとして扱うことができます。 [ShapeGroup は](/javascript/api/excel/excel.shapegroup)タイプの 1 つなので、アドインはグループを 1 つの `Shape` 図形として扱います。

次のコード サンプルは、グループ化されている 3 つの図形を示しています。 以降のコード サンプルは、図形グループが右 50 ピクセルに移動されているのを示しています。

```js
// This sample takes three previously-created shapes ("Square", "Pentagon", and "Octagon")
// and groups them into a single ShapeGroup.
Excel.run(function (context) {
    var shapes = context.workbook.worksheets.getItem("MyWorksheet").shapes;
    var square = shapes.getItem("Square");
    var pentagon = shapes.getItem("Pentagon");
    var octagon = shapes.getItem("Octagon");

    var shapeGroup = shapes.addGroup([square, pentagon, octagon]);
    shapeGroup.name = "Group";
    console.log("Shapes grouped");

    return context.sync();
}).catch(errorHandlerFunction);

// This sample moves the previously created shape group to the right by 50 pixels.
Excel.run(function (context) {
    var shapes = context.workbook.worksheets.getItem("MyWorksheet").shapes;
    var shapeGroup = sheet.shapes.getItem("Group");
    shapeGroup.incrementLeft(50);
    return context.sync();
}).catch(errorHandlerFunction);
```

> [!IMPORTANT]
> グループ内の個々の図形は `ShapeGroup.shapes` [、GroupShapeCollection](/javascript/api/excel/excel.GroupShapeCollection)型のプロパティを介して参照されます。 グループ化された後、ワークシートの図形コレクションからアクセスできなくなりました。 たとえば、ワークシートに 3 つの図形が含み、すべての図形がグループ化されている場合、ワークシートのメソッドは `shapes.getCount` カウント 1 を返します。

## <a name="export-shapes-as-images"></a>図形をイメージとしてエクスポートする

任意 `Shape` のオブジェクトをイメージに変換できます。 [Shape.getAsImage](/javascript/api/excel/excel.shape#getAsImage_format_) は base64 エンコードされた文字列を返します。 イメージの形式は、に渡される [PictureFormat](/javascript/api/excel/excel.pictureformat) 列挙型として指定されます `getAsImage` 。

```js
Excel.run(function (context) {
    var shapes = context.workbook.worksheets.getItem("MyWorksheet").shapes;
    var shape = sheet.shapes.getItem("Image");
    var stringResult = shape.getAsImage(Excel.PictureFormat.png);

    return context.sync().then(function () {
        console.log(stringResult.value);
        // Instead of logging, your add-in may use the base64-encoded string to save the image as a file or insert it in HTML.
    });
}).catch(errorHandlerFunction);
```

## <a name="delete-shapes"></a>図形を削除する

図形は、オブジェクトのメソッドを使用して `Shape` ワークシートから削除 `delete` されます。 他のメタデータは不要です。

次のコード サンプルでは **、MyWorksheet** からすべての図形を削除します。

```js
// This deletes all the shapes from "MyWorksheet".
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("MyWorksheet");
    var shapes = sheet.shapes;

    // We'll load all the shapes in the collection without loading their properties.
    shapes.load("items/$none");
    return context.sync().then(function () {
        shapes.items.forEach(function (shape) {
            shape.delete()
        });
        return context.sync();
    }).catch(errorHandlerFunction);
}).catch(errorHandlerFunction);
```

## <a name="see-also"></a>関連項目

- [Excel JavaScript API を使用した基本的なプログラミングの概念](../reference/overview/excel-add-ins-reference-overview.md)
- [Excel JavaScript API を使用してグラフを操作する](excel-add-ins-charts.md)
