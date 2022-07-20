---
title: Excel JavaScript API を使用して図形を操作する
description: Excel で図形を Excel の描画レイヤー上にある任意のオブジェクトとして定義する方法について説明します。
ms.date: 02/17/2022
ms.localizationpriority: medium
ms.openlocfilehash: 507ae05b570e7eef4f3bf5560ca47c1bfbd40f9f
ms.sourcegitcommit: df7964b6509ee6a807d754fbe895d160bc52c2d3
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/20/2022
ms.locfileid: "66889598"
---
# <a name="work-with-shapes-using-the-excel-javascript-api"></a>Excel JavaScript API を使用して図形を操作する

Excel では、図形を Excel の描画レイヤー上にある任意のオブジェクトとして定義します。 つまり、セルの外側にあるものは図形です。 この記事では、図形、線、および画像を [Shape](/javascript/api/excel/excel.shape) API および [ShapeCollection](/javascript/api/excel/excel.shapecollection) API と組み合わせて使用する方法について説明します。 [グラフについては](/javascript/api/excel/excel.chart) 、独自の記事「 [Excel JavaScript API を使用してグラフを操作する」を参照してください](excel-add-ins-charts.md)。

次の図は、温度計を形成する図形を示しています。
![Excel 図形として作成された温度計の画像。](../images/excel-shapes.png)

## <a name="create-shapes"></a>図形を作成する

図形は、ワークシートの図形コレクション (`Worksheet.shapes`) を通じて作成され、保存されます。 `ShapeCollection` には、この目的のためにいくつかの `.add*` メソッドがあります。 すべての図形には、コレクションに追加されるときに、それらの名前と ID が生成されます。 `name`これらはそれぞれプロパティと`id`プロパティです。 `name` は、アドインによって設定できるため、メソッドを使用して簡単に `ShapeCollection.getItem(name)` 取得できます。

関連付けられたメソッドを使用して、次の種類の図形が追加されます。

| Shape | Tabs.Add メソッド (Outlook フォーム スクリプト) | 署名 |
|-------|------------|-----------|
| ジオメトリックシェイプ | [addGeometricShape](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-addgeometricshape-member(1)) | `addGeometricShape(geometricShapeType: Excel.GeometricShapeType): Excel.Shape` |
| 画像 (JPEG または PNG) | [Addimage](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-addimage-member(1)) | `addImage(base64ImageString: string): Excel.Shape` |
| Line | [addLine](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-addline-member(1)) | `addLine(startLeft: number, startTop: number, endLeft: number, endTop: number, connectorType?: Excel.ConnectorType): Excel.Shape` |
| SVG | [addSvg](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-addsvg-member(1)) | `addSvg(xml: string): Excel.Shape` |
| テキスト ボックス | [addTextBox](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-addtextbox-member(1)) | `addTextBox(text?: string): Excel.Shape` |

### <a name="geometric-shapes"></a>幾何学的図形

ジオメトリ図形は 、 `ShapeCollection.addGeometricShape`. このメソッドは、 [GeometrShapeType](/javascript/api/excel/excel.geometricshapetype) 列挙型を引数として受け取ります。

次のコード サンプルでは、ワークシートの上辺と左側から 100 ピクセルの位置にある **"Square"** という名前の 150 x 150 ピクセルの四角形を作成します。

```js
// This sample creates a rectangle positioned 100 pixels from the top and left sides
// of the worksheet and is 150x150 pixels.
await Excel.run(async (context) => {
    let shapes = context.workbook.worksheets.getItem("MyWorksheet").shapes;

    let rectangle = shapes.addGeometricShape(Excel.GeometricShapeType.rectangle);
    rectangle.left = 100;
    rectangle.top = 100;
    rectangle.height = 150;
    rectangle.width = 150;
    rectangle.name = "Square";

    await context.sync();
});
```

### <a name="images"></a>画像

JPEG、PNG、SVG イメージは、図形としてワークシートに挿入できます。 このメソッドは `ShapeCollection.addImage` 、base64 でエンコードされた文字列を引数として受け取ります。 これは、文字列形式の JPEG または PNG イメージです。 `ShapeCollection.addSvg` この引数はグラフィックを定義する XML ですが、文字列も取り込みます。

次のコード サンプルは、 [FileReader](https://developer.mozilla.org/docs/Web/API/FileReader) によって文字列として読み込まれるイメージ ファイルを示しています。 文字列には、図形が作成される前に削除されたメタデータ "base64" があります。

```js
// This sample creates an image as a Shape object in the worksheet.
let myFile = document.getElementById("selectedFile");
let reader = new FileReader();

reader.onload = (event) => {
    Excel.run(function (context) {
        let startIndex = reader.result.toString().indexOf("base64,");
        let myBase64 = reader.result.toString().substr(startIndex + 7);
        let sheet = context.workbook.worksheets.getItem("MyWorksheet");
        let image = sheet.shapes.addImage(myBase64);
        image.name = "Image";
        return context.sync();
    }).catch(errorHandlerFunction);
};

// Read in the image file as a data URL.
reader.readAsDataURL(myFile.files[0]);
```

### <a name="lines"></a>Lines

行は .`ShapeCollection.addLine` このメソッドには、線の始点と終点の左余白と上余白が必要です。 また、コネクタ [型](/javascript/api/excel/excel.connectortype) 列挙型を使用して、エンドポイント間で線を制御する方法を指定します。 次のコード サンプルでは、ワークシートに直線を作成します。

```js
// This sample creates a straight line from [200,50] to [300,150] on the worksheet.
await Excel.run(async (context) => {
    let shapes = context.workbook.worksheets.getItem("MyWorksheet").shapes;
    let line = shapes.addLine(200, 50, 300, 150, Excel.ConnectorType.straight);
    line.name = "StraightLine";
    await context.sync();
});
```

線は、他の Shape オブジェクトに接続できます。 およびメソッドは `connectBeginShape` 、 `connectEndShape` 指定した接続ポイントの図形に線の始点と末尾をアタッチします。 これらのポイントの場所は図形によって異なりますが `Shape.connectionSiteCount` 、アドインが範囲外のポイントに接続しないようにするために使用できます。 線は、接続されている図形から、および`disconnectEndShape`メソッドを`disconnectBeginShape`使用して切断されます。

次のコード サンプルでは **、"MyLine"** 行を **"LeftShape" と "RightShape"** という名前の 2 つの図形に接続します。

```js
// This sample connects a line between two shapes at connection points '0' and '3'.
await Excel.run(async (context) => {
    let shapes = context.workbook.worksheets.getItem("MyWorksheet").shapes;
    let line = shapes.getItem("MyLine").line;
    line.connectBeginShape(shapes.getItem("LeftShape"), 0);
    line.connectEndShape(shapes.getItem("RightShape"), 3);
    await context.sync();
});
```

## <a name="move-and-resize-shapes"></a>図形の移動とサイズ変更

図形はワークシートの上に配置されます。 配置は、and `top` プロパティによって`left`定義されます。 これらはワークシートのそれぞれの端からの余白として機能し、[0, 0] は左上隅です。 これらは、直接設定することも、メソッドを`incrementTop`使用`incrementLeft`して現在の位置から調整することもできます。 既定の位置から回転される図形の量も、この方法で確立され、プロパティは絶対量であり、 `rotation` メソッドは既存の `incrementRotation` 回転を調整します。

他の図形に対する図形の奥行きは、プロパティによって `zorderPosition` 定義されます。 これは、[ShapeZOrder](/javascript/api/excel/excel.shapezorder) を`setZOrder`受け取るメソッドを使用して設定されます。 `setZOrder` は、他の図形に対する現在の図形の順序を調整します。

アドインには、図形の高さと幅を変更するためのいくつかのオプションがあります。 または`width`プロパティを設定すると、`height`他のディメンションを変更せずに、指定したディメンションが変更されます。 `scaleHeight` `scaleWidth`現在または元のサイズ (指定された [ShapeScaleType](/javascript/api/excel/excel.shapescaletype) の値に基づいて) に対して、図形のそれぞれの寸法を調整します。 省略可能な [ShapeScaleFrom](/javascript/api/excel/excel.shapescalefrom) パラメーターは、図形のスケールの位置 (左上隅、中央、または右下隅) を指定します。 プロパティの `lockAspectRatio` 場合、スケール メソッドは `true`、もう一方の寸法も調整することで、図形の現在の縦横比を維持します。

> [!NOTE]
> プロパティと`width`プロパティへの直接変更は`height`、プロパティの値に関係なく、そのプロパティにのみ影響します`lockAspectRatio`。

次のコード サンプルは、元のサイズの 1.25 倍にスケーリングされ、30 度回転する図形を示しています。

```js
// In this sample, the shape "Octagon" is rotated 30 degrees clockwise
// and scaled 25% larger, with the upper-left corner remaining in place.
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("MyWorksheet");

    let shape = sheet.shapes.getItem("Octagon");
    shape.incrementRotation(30);
    shape.lockAspectRatio = true;
    shape.scaleWidth(
        1.25,
        Excel.ShapeScaleType.currentSize,
        Excel.ShapeScaleFrom.scaleFromTopLeft);

    await context.sync();
});
```

## <a name="text-in-shapes"></a>図形内のテキスト

ジオメトリ図形にはテキストを含めることができます。 図形には [TextFrame](/javascript/api/excel/excel.textframe)`textFrame` 型のプロパティがあります。 オブジェクトは `TextFrame` 、テキスト表示オプション (余白やテキスト オーバーフローなど) を管理します。 `TextFrame.textRange` は、テキストの内容とフォント設定を含む [TextRange](/javascript/api/excel/excel.textrange) オブジェクトです。

次のコード サンプルでは、"Shape Text" というテキストを含む "Wave" という名前の幾何学的図形を作成します。 また、図形とテキストの色を調整し、テキストの水平方向の配置を中央に設定します。

```js
// This sample creates a light-blue wave shape and adds the purple text "Shape text" to the center.
await Excel.run(async (context) => {
    let shapes = context.workbook.worksheets.getItem("MyWorksheet").shapes;
    let wave = shapes.addGeometricShape(Excel.GeometricShapeType.wave);
    wave.left = 100;
    wave.top = 400;
    wave.height = 50;
    wave.width = 150;

    wave.name = "Wave";
    wave.fill.setSolidColor("lightblue");

    wave.textFrame.textRange.text = "Shape text";
    wave.textFrame.textRange.font.color = "purple";
    wave.textFrame.horizontalAlignment = Excel.ShapeTextHorizontalAlignment.center;

    await context.sync();
});
```

`ShapeCollection`メソッドは`addTextBox`、`GeometricShape`白い背景と黒のテキストを持つ型`Rectangle`を作成します。 これは、[**挿入**] タブの Excel の **[テキスト ボックス]** ボタンで作成したものと同じです。 `addTextBox`のテキストを設定する文字列引数を`TextRange`受け取ります。

次のコード サンプルは、テキスト "Hello!"" を含むテキスト ボックスの作成を示しています。

```js
// This sample creates a text box with the text "Hello!" and sizes it appropriately.
await Excel.run(async (context) => {
    let shapes = context.workbook.worksheets.getItem("MyWorksheet").shapes;
    let textbox = shapes.addTextBox("Hello!");
    textbox.left = 100;
    textbox.top = 100;
    textbox.height = 20;
    textbox.width = 45;
    textbox.name = "Textbox";
    await context.sync();
});
```

## <a name="shape-groups"></a>図形グループ

図形はグループ化できます。 これにより、ユーザーは、それらを配置、サイズ設定、およびその他の関連タスク用の 1 つのエンティティとして扱うことができます。 [ShapeGroup](/javascript/api/excel/excel.shapegroup) は一種であるため`Shape`、アドインはグループを 1 つの図形として扱います。

次のコード サンプルは、3 つの図形がグループ化されていることを示しています。 その後のコード サンプルでは、図形グループが右側の 50 ピクセルに移動されていることを示しています。

```js
// This sample takes three previously-created shapes ("Square", "Pentagon", and "Octagon")
// and groups them into a single ShapeGroup.
await Excel.run(async (context) => {
    let shapes = context.workbook.worksheets.getItem("MyWorksheet").shapes;
    let square = shapes.getItem("Square");
    let pentagon = shapes.getItem("Pentagon");
    let octagon = shapes.getItem("Octagon");

    let shapeGroup = shapes.addGroup([square, pentagon, octagon]);
    shapeGroup.name = "Group";
    console.log("Shapes grouped");

    await context.sync();
});

// This sample moves the previously created shape group to the right by 50 pixels.
await Excel.run(async (context) => {
    let shapes = context.workbook.worksheets.getItem("MyWorksheet").shapes;
    let shapeGroup = shapes.getItem("Group");
    shapeGroup.incrementLeft(50);
    await context.sync();
});
```

> [!IMPORTANT]
> グループ内の個々の図形は、[GroupShapeCollection](/javascript/api/excel/excel.groupshapecollection) 型のプロパティを介して`ShapeGroup.shapes`参照されます。 グループ化された後は、ワークシートの図形コレクションからアクセスできなくなります。 たとえば、ワークシートに 3 つの図形があり、それらがすべてグループ化されている場合、ワークシートの `shapes.getCount` メソッドはカウント 1 を返します。

## <a name="export-shapes-as-images"></a>図形を画像としてエクスポートする

任意の `Shape` オブジェクトをイメージに変換できます。 [Shape.getAsImage は](/javascript/api/excel/excel.shape#excel-excel-shape-getasimage-member(1)) base64 でエンコードされた文字列を返します。 イメージの形式は、渡される `getAsImage`[PictureFormat](/javascript/api/excel/excel.pictureformat) 列挙型として指定されます。

```js
await Excel.run(async (context) => {
    let shapes = context.workbook.worksheets.getItem("MyWorksheet").shapes;
    let shape = shapes.getItem("Image");
    let stringResult = shape.getAsImage(Excel.PictureFormat.png);

    await context.sync();

    console.log(stringResult.value);
    // Instead of logging, your add-in may use the base64-encoded string to save the image as a file or insert it in HTML.
});
```

## <a name="delete-shapes"></a>図形を削除する

図形は、オブジェクト`delete`のメソッドを使用して`Shape`ワークシートから削除されます。 他のメタデータは必要ありません。

次のコード サンプルでは、 **MyWorksheet** からすべての図形を削除します。

```js
// This deletes all the shapes from "MyWorksheet".
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("MyWorksheet");
    let shapes = sheet.shapes;

    // We'll load all the shapes in the collection without loading their properties.
    shapes.load("items/$none");
    await context.sync();

    shapes.items.forEach(function (shape) {
        shape.delete();
    });
    
    await context.sync();
});
```

## <a name="see-also"></a>関連項目

- [Excel JavaScript API を使用した基本的なプログラミングの概念](../reference/overview/excel-add-ins-reference-overview.md)
- [Excel JavaScript API を使用してグラフを操作する](excel-add-ins-charts.md)
