---
title: Excel JavaScript API を使用して図形を操作する
description: Excel の描画レイヤー上にある任意のオブジェクトとして、Excel によって図形が定義される方法について説明します。
ms.date: 01/14/2020
localization_priority: Normal
ms.openlocfilehash: 7b9a4dba02e28187eeb0f932e245489ca61fcbcc
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/08/2020
ms.locfileid: "44609742"
---
# <a name="work-with-shapes-using-the-excel-javascript-api"></a>Excel JavaScript API を使用して図形を操作する

Excel では、図形は Excel の描画層にある任意のオブジェクトとして定義されます。 つまり、セルの外部にあるものは図形です。 この記事では、[図形](/javascript/api/excel/excel.shape)および shapes [ecollection](/javascript/api/excel/excel.shapecollection) api と組み合わせて、ジオメトリック図形、線、およびイメージを使用する方法について説明します。 [グラフ](/javascript/api/excel/excel.chart)については、「 [Excel JavaScript API を使用してグラフを操作](excel-add-ins-charts.md)する」で説明されています。

次の図は、温度計を形成する図形を示しています。
![Excel 図形として作成された温度計のイメージ](../images/excel-shapes.png)

## <a name="create-shapes"></a>図形を作成する

図形は、ワークシートの shape コレクション () を使用して作成され、格納され `Worksheet.shapes` ます。 `ShapeCollection`には `.add*` 、この目的のためにいくつかの方法があります。 すべての図形には、コレクションに追加されたときに名前と Id が生成されます。 これらは、 `name` および `id` プロパティです。 `name`アドインで設定して、メソッドを使用して簡単に取得することができ `ShapeCollection.getItem(name)` ます。

次の種類の図形は、関連付けられているメソッドを使用して追加されます。

| Shape | Tabs.Add メソッド (Outlook フォーム スクリプト) | 署名 |
|-------|------------|-----------|
| 幾何学的図形 | [addGeometricShape](/javascript/api/excel/excel.shapecollection#addgeometricshape-geometricshapetype-) | `addGeometricShape(geometricShapeType: Excel.GeometricShapeType): Excel.Shape` |
| 画像 (JPEG または PNG のいずれか) | [addImage](/javascript/api/excel/excel.shapecollection#addimage-base64imagestring-) | `addImage(base64ImageString: string): Excel.Shape` |
| 線 | [addLine](/javascript/api/excel/excel.shapecollection#addline-startleft--starttop--endleft--endtop--connectortype-) | `addLine(startLeft: number, startTop: number, endLeft: number, endTop: number, connectorType?: Excel.ConnectorType): Excel.Shape` |
| SVG | [addSvg](/javascript/api/excel/excel.shapecollection#addsvg-xml-) | `addSvg(xml: string): Excel.Shape` |
| テキスト ボックス | [addTextBox](/javascript/api/excel/excel.shapecollection#addtextbox-text-) | `addTextBox(text?: string): Excel.Shape` |

### <a name="geometric-shapes"></a>幾何学的な図形

ジオメトリック図形が作成され `ShapeCollection.addGeometricShape` ます。 このメソッドは、 [GeometricShapeType](/javascript/api/excel/excel.geometricshapetype) enum を引数として受け取ります。

次のコードサンプルでは、ワークシートの上端と左端から100ピクセルに配置された、 **"Square"** という名前の150x150 ピクセルの四角形を作成します。

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

JPEG、PNG、SVG の画像は、図形としてワークシートに挿入できます。 メソッドは、 `ShapeCollection.addImage` base64 でエンコードされた文字列を引数として受け取ります。 これは、文字列形式の JPEG または PNG 画像のいずれかです。 `ShapeCollection.addSvg`も文字列で受け取りますが、この引数はグラフィックを定義する XML です。

次のコードサンプルは、 [FileReader](https://developer.mozilla.org/docs/Web/API/FileReader)によって、文字列としてロードされているイメージファイルを示しています。 文字列には、図形が作成される前に、メタデータ "base64" が削除されています。

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

行はに作成され `ShapeCollection.addLine` ます。 このメソッドには、行の開始点と終了点の左余白と上余白が必要です。 また、 [ConnectorType](/javascript/api/excel/excel.connectortype)列挙を取得して、エンドポイント間の行の contorts 方法を指定します。 次のコードサンプルでは、ワークシートに直線を作成します。

```js
// This sample creates a straight line from [200,50] to [300,150] on the worksheet
Excel.run(function (context) {
    var shapes = context.workbook.worksheets.getItem("MyWorksheet").shapes;
    var line = shapes.addLine(200, 50, 300, 150, Excel.ConnectorType.straight);
    line.name = "StraightLine";
    return context.sync();
}).catch(errorHandlerFunction);
```

線は、他の Shape オブジェクトに接続することができます。 `connectBeginShape`メソッドと `connectEndShape` メソッドは、指定された接続ポイントにある図形に対して、線の始点と終点を接続します。 これらのポイントの位置は図形によって異なりますが、を使用すると、 `Shape.connectionSiteCount` アドインが範囲外のポイントに接続されないようにすることができます。 およびメソッドを使用して、接続されているすべての図形から線が切断され `disconnectBeginShape` `disconnectEndShape` ます。

次のコードサンプルでは、" **Myline"** 行を **"l shape"** と **"直角図形"** という名前の2つの図形に接続します。

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

ワークシートの一番上にある図形。 これらの配置は、およびプロパティによって定義され `left` `top` ます。 これらは、ワークシートの各エッジの余白として機能し、[0, 0] が左上隅になります。 これらは、およびメソッドを使用して、現在の位置から直接設定または調整することができ `incrementLeft` `incrementTop` ます。 既定の位置から図形を回転させる度合いは、この方法で設定します。この方法では、プロパティが絶対量で、既存の回転を調整するメソッドも使用され `rotation` `incrementRotation` ます。

他の図形を基準とした図形の深さは、プロパティによって定義され `zorderPosition` ます。 これはメソッドを使用して設定されます。このメソッドは、このメソッドを使用し `setZOrder` ます。 [ShapeZOrder](/javascript/api/excel/excel.shapezorder) `setZOrder`他の図形を基準に現在の図形の順序を調整します。

アドインには、図形の高さと幅を変更するためのいくつかのオプションがあります。 またはプロパティのいずれかを設定すると、 `height` `width` 他の次元を変更せずに、指定した次元が変更されます。 `scaleHeight`を指定して、 `scaleWidth` 現在のサイズまたは元のサイズを基準にして図形のそれぞれの寸法を調整します (提供される[ShapeScaleType](/javascript/api/excel/excel.shapescaletype)の値に基づきます)。 省略可能な[ShapeScaleFrom](/javascript/api/excel/excel.shapescalefrom)パラメーターは、図形を拡大または縮小する位置 (左上隅、中央、または右下隅) を指定します。 プロパティに `lockAspectRatio` **true**が設定されている場合、scale メソッドは、他の次元も調整して、図形の現在の縦横比を維持します。

> [!NOTE]
> プロパティに対する直接の変更は、プロパティの `height` `width` 値に関係なく、そのプロパティにのみ影響し `lockAspectRatio` ます。

次のコードサンプルでは、元のサイズに1.25 倍に拡大または縮小された図形を表示します。

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

幾何学的な図形にはテキストを含めることができます。 図形には、 `textFrame` [TextFrame](/javascript/api/excel/excel.textframe)型のプロパティがあります。 オブジェクトは、 `TextFrame` テキスト表示オプション (余白、テキストオーバーフローなど) を管理します。 `TextFrame.textRange`は、テキストの内容とフォントの設定を含む[TextRange](/javascript/api/excel/excel.textrange)オブジェクトです。

次のコードサンプルでは、テキスト "Shape Text" を使用して "Wave" という名前のジオメトリック図形を作成します。 また、図形とテキストの色を調整するだけでなく、テキストの水平方向の配置を中央に設定します。

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

`addTextBox`のメソッドは、 `ShapeCollection` `GeometricShape` 白の `Rectangle` 背景と黒のテキストを使用して、型を作成します。 これは、[挿入] タブの [Excel の**テキストボックス**での作成**Insert**方法] ボタンと同じです。 `addTextBox`文字列引数を受け取り、のテキストを設定し `TextRange` ます。

次のコードサンプルは、"Hello!" というテキストを含むテキストボックスを作成する方法を示しています。

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

図形は一緒にグループ化できます。 これにより、ユーザーは、配置、サイズ変更、およびその他の関連タスクのために1つのエンティティとして扱うことができます。 図形[グループ](/javascript/api/excel/excel.shapegroup)はの種類である `Shape` ため、アドインでグループを1つの図形として扱うことができます。

次のコードサンプルでは、グループ化された3つの図形を示します。 次のコードサンプルでは、図形グループを50ピクセル右に移動していることを示します。

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
> グループ内の個々の図形は、 `ShapeGroup.shapes` 種類が[Groupshapecollection](/javascript/api/excel/excel.GroupShapeCollection)であるプロパティを介して参照されます。 グループ化された後は、ワークシートの shape コレクションからアクセスできなくなります。 たとえば、ワークシートに3つの図形があり、すべてが一緒にグループ化されている場合、ワークシートの `shapes.getCount` メソッドはカウントを1とします。

## <a name="export-shapes-as-images"></a>図形を画像としてエクスポートする

任意の `Shape` オブジェクトをイメージに変換できます。 [GetAsImage](/javascript/api/excel/excel.shape#getasimage-format-)は、base64 でエンコードされた文字列を返します。 画像の形式は、に渡される図[形式](/javascript/api/excel/excel.pictureformat)の列挙体として指定され `getAsImage` ます。

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

図形は、オブジェクトのメソッドを使用してワークシートから削除され `Shape` `delete` ます。 その他のメタデータは必要ありません。

次のコードサンプルでは、 **Myworksheet**からすべての図形を削除します。

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
