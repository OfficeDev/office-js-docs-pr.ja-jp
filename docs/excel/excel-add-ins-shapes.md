---
title: JavaScript API を使用して図形をExcelする
description: 図形をExcel図面レイヤーに配置するオブジェクトとして定義する方法についてExcel。
ms.date: 02/17/2022
ms.localizationpriority: medium
ms.openlocfilehash: 1e268f32f7388a1992d46c53bbb8077d605e9fb7
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/23/2022
ms.locfileid: "63745527"
---
# <a name="work-with-shapes-using-the-excel-javascript-api"></a>JavaScript API を使用して図形をExcelする

Excel図形は、図形の描画レイヤーに配置されるオブジェクトとして定義Excel。 つまり、セルの外側にあるものは図形です。 この記事では、図形および [ShapeCollection](/javascript/api/excel/excel.shapecollection) API と組み合わせて幾何学的図形[](/javascript/api/excel/excel.shape)、線、および画像を使用する方法について説明します。 [グラフ](/javascript/api/excel/excel.chart)については、独自の記事「[JavaScript API](excel-add-ins-charts.md) を使用してグラフをExcelします。

次の図は、体温計を形成する図形を示しています。
![図形として作成された体温計Excelします。](../images/excel-shapes.png)

## <a name="create-shapes"></a>図形を作成する

図形は、ワークシートの図形コレクション () を通じて作成され、保存されます`Worksheet.shapes`。 `ShapeCollection` この目的のために `.add*` いくつかのメソッドがあります。 すべての図形には、コレクションに追加するときに名前と ID が生成されます。 これらは、それぞれプロパティ`name``id`とプロパティです。 `name` メソッドを使用して簡単に取得できるようアドインで設定 `ShapeCollection.getItem(name)` できます。

次の種類の図形は、関連付けられたメソッドを使用して追加されます。

| Shape | Tabs.Add メソッド (Outlook フォーム スクリプト) | 署名 |
|-------|------------|-----------|
| ジオメトリック シェイプ | [addGeometricShape](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-addgeometricshape-member(1)) | `addGeometricShape(geometricShapeType: Excel.GeometricShapeType): Excel.Shape` |
| 画像 (JPEG または PNG) | [addImage](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-addimage-member(1)) | `addImage(base64ImageString: string): Excel.Shape` |
| Line | [addLine](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-addline-member(1)) | `addLine(startLeft: number, startTop: number, endLeft: number, endTop: number, connectorType?: Excel.ConnectorType): Excel.Shape` |
| SVG | [addSvg](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-addsvg-member(1)) | `addSvg(xml: string): Excel.Shape` |
| テキスト ボックス | [addTextBox](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-addtextbox-member(1)) | `addTextBox(text?: string): Excel.Shape` |

### <a name="geometric-shapes"></a>幾何学的図形

で幾何学的な図形が作成されます `ShapeCollection.addGeometricShape`。 このメソッドは、 [引数として GeometricShapeType](/javascript/api/excel/excel.geometricshapetype) 列挙型を受け取ります。

次のコード サンプルでは、ワークシートの上辺と左側から 100 ピクセルの位置にある **"Square"** という名前の 150x150 ピクセルの四角形を作成します。

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

JPEG、PNG、SVG 画像は、ワークシートに図形として挿入できます。 メソッド `ShapeCollection.addImage` は引数として base64 でエンコードされた文字列を受け取ります。 これは、文字列形式の JPEG または PNG イメージのいずれかです。 `ShapeCollection.addSvg` この引数はグラフィックを定義する XML ですが、文字列も取り込まれます。

次のコード サンプルは、FileReader が文字列として読み込 [むイメージ ファイル](https://developer.mozilla.org/docs/Web/API/FileReader) を示しています。 この文字列には、図形が作成される前にメタデータ "base64" が削除されます。

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

行がで作成されます `ShapeCollection.addLine`。 このメソッドには、行の開始点と終了点の左余白と上余白が必要です。 また、 [ConnectorType 列挙型を使用](/javascript/api/excel/excel.connectortype) して、エンドポイント間の行のコントルト方法を指定します。 次のコード サンプルでは、ワークシートに直線を作成します。

```js
// This sample creates a straight line from [200,50] to [300,150] on the worksheet.
await Excel.run(async (context) => {
    let shapes = context.workbook.worksheets.getItem("MyWorksheet").shapes;
    let line = shapes.addLine(200, 50, 300, 150, Excel.ConnectorType.straight);
    line.name = "StraightLine";
    await context.sync();
});
```

線は他の Shape オブジェクトに接続できます。 and `connectBeginShape` メソッド `connectEndShape` は、指定した接続ポイントの図形に線の開始位置と終了位置をアタッチします。 これらのポイントの `Shape.connectionSiteCount` 位置は図形によって異なりますが、アドインが境界外のポイントに接続しない場合に使用できます。 線は、and メソッドを使用して、接続されている図形から`disconnectBeginShape``disconnectEndShape`切断されます。

次のコード サンプルでは、 **"MyLine"** 行を **"LeftShape" と "RightShape" という名前** の 2 つの図形 **に接続します**。

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

図形はワークシートの上に表示されます。 配置は and プロパティによって `left` 定義 `top` されます。 これらはワークシートのそれぞれのエッジの余白として機能し、[0, 0] は左上隅になります。 これらは、and メソッドを使用して、直接設定するか、現在の位置から `incrementLeft` 調整 `incrementTop` できます。 既定の位置から`rotation``incrementRotation`回転する図形の量も、この方法で確立され、プロパティは絶対量であり、メソッドは既存の回転を調整します。

他の図形に対する図形の深さは、プロパティによって定義 `zorderPosition` されます。 これは、ShapeZOrder `setZOrder` を受け取るメソッドを [使用して設定されます](/javascript/api/excel/excel.shapezorder)。 `setZOrder` 他の図形を基準に現在の図形の順序を調整します。

アドインには、図形の高さと幅を変更するためのいくつかのオプションがあります。 またはプロパティを設定 `height` すると `width` 、他のディメンションを変更せずに指定したディメンションが変更されます。 現在 `scaleHeight` のサイズ `scaleWidth` または元のサイズ (指定された [ShapeScaleType](/javascript/api/excel/excel.shapescaletype) の値に基づいて) を基準に図形のそれぞれの寸法を調整します。 省略可能 [な ShapeScaleFrom](/javascript/api/excel/excel.shapescalefrom) パラメーターは、図形のスケールの場所 (左上隅、中央、右下隅) を指定します。 プロパティが `lockAspectRatio` true の **場合**、スケール メソッドは他のディメンションも調整することで、図形の現在の縦横比を維持します。

> [!NOTE]
> プロパティへの直接の変更 `height` は `width` 、プロパティの値に関係なく `lockAspectRatio` 、そのプロパティにのみ影響します。

次のコード サンプルは、元のサイズの 1.25 倍に拡大縮小され、30 度回転された図形を示しています。

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

幾何学的図形にはテキストを含めできます。 図形には、`textFrame`[TextFrame 型のプロパティがあります](/javascript/api/excel/excel.textframe)。 オブジェクト `TextFrame` は、テキスト表示オプション (余白やテキスト オーバーフローなど) を管理します。 `TextFrame.textRange` は、 [テキストコンテンツと](/javascript/api/excel/excel.textrange) フォント設定を持つ TextRange オブジェクトです。

次のコード サンプルでは、"Shape Text" というテキストを持つ "Wave" という名前の幾何学的な図形を作成します。 また、図形とテキストの色を調整し、テキストの水平方向の配置を中央に設定します。

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

この `addTextBox` メソッドは、 `ShapeCollection` 白い背景と `GeometricShape` 黒い `Rectangle` テキストを持つ型を作成します。 これは、[挿入] タブ`addTextBox`の Excel の [テキスト ボックス] ボタンによって作成される操作と同じです。文字列引数を使用して、 のテキストを設定します。`TextRange`

次のコード サンプルは、テキスト "Hello!" を含むテキスト ボックスの作成を示しています。

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

図形はグループ化できます。 これにより、ユーザーはそれらを配置、サイズ変更、その他の関連タスク用の 1 つのエンティティとして扱うことができます。 [ShapeGroup は](/javascript/api/excel/excel.shapegroup)タイプの 1 つ`Shape`なので、アドインはグループを 1 つの図形として扱います。

次のコード サンプルは、グループ化されている 3 つの図形を示しています。 以降のコード サンプルは、図形グループが右 50 ピクセルに移動されているのを示しています。

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
> グループ内の個々の図形は、`ShapeGroup.shapes`[GroupShapeCollection 型のプロパティを介して参照されます](/javascript/api/excel/excel.groupshapecollection)。 グループ化された後、ワークシートの図形コレクションからアクセスできなくなりました。 たとえば、ワークシートに 3 `shapes.getCount` つの図形が含み、すべての図形がグループ化されている場合、ワークシートのメソッドはカウント 1 を返します。

## <a name="export-shapes-as-images"></a>図形をイメージとしてエクスポートする

任意 `Shape` のオブジェクトをイメージに変換できます。 [Shape.getAsImage](/javascript/api/excel/excel.shape#excel-excel-shape-getasimage-member(1)) は base64 エンコードされた文字列を返します。 イメージの形式は、に渡される [PictureFormat](/javascript/api/excel/excel.pictureformat) 列挙型として指定されます `getAsImage`。

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

図形は、オブジェクトのメソッドを使用してワークシート `Shape` から削除 `delete` されます。 他のメタデータは不要です。

次のコード サンプルでは、 **MyWorksheet からすべての図形を削除します**。

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
