---
title: PowerPoint JavaScript API を使用して図形を操作する
description: PowerPoint スライドで図形を追加、削除、書式設定する方法について説明します。
ms.date: 06/13/2022
ms.localizationpriority: medium
ms.openlocfilehash: 7f314cfebb26450e79dbabe1e65ac9e4c8fe9799
ms.sourcegitcommit: 4f19f645c6c1e85b16014a342e5058989fe9a3d2
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/15/2022
ms.locfileid: "66091105"
---
# <a name="work-with-shapes-using-the-powerpoint-javascript-api"></a>PowerPoint JavaScript API を使用して図形を操作する

この記事では、図形、線、テキスト ボックスを [Shape](/javascript/api/powerpoint/powerpoint.shape) API および [ShapeCollection](/javascript/api/powerpoint/powerpoint.shapecollection) API と組み合わせて使用する方法について説明します。

## <a name="create-shapes"></a>図形を作成する

図形はスライドの図形コレクション (`slide.shapes`) を介して作成され、格納されます。 `ShapeCollection` には、この目的のためにいくつかの `.add*` メソッドがあります。 すべての図形には、コレクションに追加されるときに、それらの名前と ID が生成されます。 `name`これらはそれぞれプロパティと`id`プロパティです。 `name` はアドインで設定できます。

### <a name="geometric-shapes"></a>幾何学的図形

のオーバーロードの 1 つを使用して、ジオメトリ図形が作成されます `ShapeCollection.addGeometricShape`。 最初のパラメーターは [、GeometrShapeType](/javascript/api/powerpoint/powerpoint.geometricshapetype) 列挙型か、列挙型の値のいずれかに相当する文字列です。 [ShapeAddOptions](/javascript/api/powerpoint/powerpoint.shapeaddoptions) 型の省略可能な 2 番目のパラメーターがあり、図形の初期サイズとその位置をスライドの上辺と左側に対する相対位置 (ポイント単位で測定) で指定できます。 または、図形の作成後にこれらのプロパティを設定することもできます。

次のコード サンプルでは、スライドの上辺と左側から 100 ポイント配置された **"Square"** という名前の四角形を作成します。 このメソッドはオブジェクトを `Shape` 返します。

```js
// This sample creates a rectangle positioned 100 points from the top and left sides
// of the slide and is 150x150 points. The shape is put on the first slide.
await PowerPoint.run(async (context) => {
    const shapes = context.presentation.slides.getItemAt(0).shapes;
    const rectangle = shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle);
    rectangle.left = 100;
    rectangle.top = 100;
    rectangle.height = 150;
    rectangle.width = 150;
    rectangle.name = "Square";
    await context.sync();
});
```

### <a name="lines"></a>Lines

のオーバーロードの 1 つを使用して行が作成されます `ShapeCollection.addLine`。 最初のパラメーターは、 [ConnectorType](/javascript/api/powerpoint/powerpoint.connectortype) 列挙型か、またはエンドポイント間での線の並べ替え方法を指定する列挙型の値のいずれかに相当する文字列です。 線の始点と終点を指定できる [ShapeAddOptions](/javascript/api/powerpoint/powerpoint.shapeaddoptions) 型の省略可能な 2 番目のパラメーターがあります。 または、図形の作成後にこれらのプロパティを設定することもできます。 このメソッドはオブジェクトを `Shape` 返します。

> [!NOTE]
> 図形が線の場合、およびオブジェクトの`Shape`プロパティと`ShapeAddOptions``left`プロパティは、`top`スライドの上端と左端に対する線の始点を指定します。 およびプロパティは`height`、`width`*始点に対する線* の終点を指定します。 そのため、スライドの上端と左端に対する終点は () で (`top``width` + `height``left` + ) です。 すべてのプロパティの測定単位はポイントであり、負の値は許可されます。

次のコード サンプルは、スライドに直線を作成します。

```js
// This sample creates a straight line on the first slide.
await PowerPoint.run(async (context) => {
    const shapes = context.presentation.slides.getItemAt(0).shapes;
    const line = shapes.addLine(Excel.ConnectorType.straight, {left: 200, top: 50, height: 300, width: 150});
    line.name = "StraightLine";
    await context.sync();
});
```

### <a name="text-boxes"></a>テキスト ボックス

[addTextBox](/javascript/api/powerpoint/powerpoint.shapecollection#powerpoint-powerpoint-shapecollection-addtextbox-member(1)) メソッドを使用してテキスト ボックスが作成されます。 最初のパラメーターは、最初にボックスに表示されるテキストです。 [ShapeAddOptions](/javascript/api/powerpoint/powerpoint.shapeaddoptions) 型の省略可能な 2 番目のパラメーターがあり、テキスト ボックスの初期サイズとそのスライドの上辺と左側を基準とした位置を指定できます。 または、図形の作成後にこれらのプロパティを設定することもできます。

次のコード サンプルは、最初のスライドにテキスト ボックスを作成する方法を示しています。

```js
// This sample creates a text box with the text "Hello!" and sizes it appropriately.
await PowerPoint.run(async (context) => {
    const shapes = context.presentation.slides.getItemAt(0).shapes;
    const textbox = shapes.addTextBox("Hello!");
    textbox.left = 100;
    textbox.top = 100;
    textbox.height = 300;
    textbox.width = 450;
    textbox.name = "Textbox";
    await context.sync();
});
```

## <a name="move-and-resize-shapes"></a>図形の移動とサイズ変更

図形はスライドの上に配置されます。 それらの配置は、プロパティと`top`プロパティによって定義されます`left`。 これらは、スライドのそれぞれの端からの余白として機能し、ポイント単位で`left: 0``top: 0`測定され、左上隅になります。 図形のサイズは、プロパティと`width`プロパティによって`height`指定されます。 コードでは、これらのプロパティをリセットすることで、図形を移動またはサイズ変更できます。 (これらのプロパティは、図形が線の場合は若干異なる意味を持ちます。 [行を](#lines)参照してください。)

## <a name="text-in-shapes"></a>図形内のテキスト

ジオメトリ図形にはテキストを含めることができます。 図形には [TextFrame](/javascript/api/powerpoint/powerpoint.textframe)`textFrame` 型のプロパティがあります。 オブジェクトは `TextFrame` 、テキスト表示オプション (余白やテキスト オーバーフローなど) を管理します。 `TextFrame.textRange` は、テキストの内容とフォント設定を含む [TextRange](/javascript/api/powerpoint/powerpoint.textrange) オブジェクトです。

次のコード サンプルでは、"Shape text" というテキストを含む **"中かっこ"** という名前のジオメトリ図形を作成 **します**。 また、図形とテキストの色を調整し、テキストの垂直方向の配置を中央に設定します。

```js
// This sample creates a light blue rectangle with braces ("{}") on the left and right ends
// and adds the purple text "Shape text" to the center.
await PowerPoint.run(async (context) => {
    const shapes = context.presentation.slides.getItemAt(0).shapes;
    const braces = shapes.addGeometricShape(PowerPoint.GeometricShapeType.bracePair);
    braces.left = 100;
    braces.top = 400;
    braces.height = 50;
    braces.width = 150;
    braces.name = "Braces";
    braces.fill.setSolidColor("lightblue");
    braces.textFrame.textRange.text = "Shape text";
    braces.textFrame.textRange.font.color = "purple";
    braces.textFrame.verticalAlignment = PowerPoint.TextVerticalAlignment.middleCentered;
    await context.sync();
});
```

## <a name="delete-shapes"></a>図形を削除する

図形は、オブジェクト`delete`のメソッドを使用して`Shape`スライドから削除されます。

次のコード サンプルは、図形を削除する方法を示しています。

```js
await PowerPoint.run(async (context) => {
    // Delete all shapes from the first slide.
    const sheet = context.presentation.slides.getItemAt(0);
    const shapes = sheet.shapes;

    // Load all the shapes in the collection without loading their properties.
    shapes.load("items/$none");
    await context.sync();
        
    shapes.items.forEach(function (shape) {
        shape.delete();
    });
    await context.sync();
});
```
