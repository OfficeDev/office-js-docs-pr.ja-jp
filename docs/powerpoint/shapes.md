---
title: JavaScript API を使用して図形をPowerPointする
description: スライドに図形を追加、削除、書式設定するPowerPointします。
ms.date: 02/22/2022
ms.localizationpriority: medium
ms.openlocfilehash: 2c7eb7a1770f807878320369951faa7d0ddc873c
ms.sourcegitcommit: 7b6ee73fa70b8e0ff45c68675dd26dd7a7b8c3e9
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/08/2022
ms.locfileid: "63340485"
---
# <a name="work-with-shapes-using-the-powerpoint-javascript-api-preview"></a>JavaScript API を使用して図形をPowerPointする (プレビュー)

この記事では、図形および [ShapeCollection](/javascript/api/powerpoint/powerpoint.shapecollection) API と組み合わせて、幾何学的図形、[](/javascript/api/powerpoint/powerpoint.shape)線、およびテキスト ボックスを使用する方法について説明します。

[!INCLUDE [Information about using preview APIs](../includes/using-preview-apis-host.md)]

## <a name="create-shapes"></a>図形を作成する

図形はスライドの図形コレクション () で作成され、格納されます`slide.shapes`。 `ShapeCollection` この目的のために `.add*` いくつかのメソッドがあります。 すべての図形には、コレクションに追加するときに名前と ID が生成されます。 これらは、それぞれプロパティ`name``id`とプロパティです。 `name` をアドインで設定できます。

### <a name="geometric-shapes"></a>幾何学的図形

幾何学的な図形は、 のオーバーロードの 1 つで作成されます `ShapeCollection.addGeometricShape`。 最初のパラメーターは [、GeometricShapeType](/javascript/api/powerpoint/powerpoint.geometricshapetype) 列挙型か、列挙型の値のいずれかと同等の文字列です。 [ShapeAddOptions](/javascript/api/powerpoint/powerpoint.shapeaddoptions) 型のオプションの 2 番目のパラメーターを使用して、スライドの上辺と左側を基準に図形の初期サイズとその位置をポイント単位で指定できます。 または、図形の作成後にこれらのプロパティを設定できます。

次のコード サンプルでは、スライドの上辺と左側から 100 ポイントの位置にある " **Square"** という名前の四角形を作成します。 メソッドはオブジェクトを返 `Shape` します。

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

行は、 のオーバーロードの 1 つで作成されます `ShapeCollection.addLine`。 最初のパラメーターは [、ConnectorType](/javascript/api/powerpoint/powerpoint.connectortype) 列挙型か、列挙型の値のいずれかと同等の文字列で、エンドポイント間の行のコントルト方法を指定します。 線の開始点と終了点を指定できる [、ShapeAddOptions](/javascript/api/powerpoint/powerpoint.shapeaddoptions) 型のオプションの 2 番目のパラメーターがあります。 または、図形の作成後にこれらのプロパティを設定できます。 メソッドはオブジェクトを返 `Shape` します。

> [!NOTE]
> 図形が線の場合 `top` `left` `Shape` `ShapeAddOptions` 、and オブジェクトのプロパティは、スライドの上端と左端を基準に線の開始点を指定します。 and `height` プロパティ `width` は、開始点を基準に線 *の端点を指定します*。 したがって、スライドの上端と左端を基準にした終了点は (`top` + `height`) です`left` + `width`。 すべてのプロパティの測定単位はポイントであり、負の値を使用できます。

次のコード サンプルでは、スライドに直線を作成します。

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

addTextBox メソッドを使用して [テキスト ボックスが作成](/javascript/api/powerpoint/powerpoint.shapecollection#powerpoint-powerpoint-shapecollection-addtextbox-member(1)) されます。 最初のパラメーターは、最初にボックスに表示されるテキストです。 [ShapeAddOptions](/javascript/api/powerpoint/powerpoint.shapeaddoptions) 型のオプションの 2 番目のパラメーターを使用して、テキスト ボックスの初期サイズと、スライドの上辺と左側の位置を指定できます。 または、図形の作成後にこれらのプロパティを設定できます。

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

図形はスライドの上に座っています。 それらの配置は、and プロパティによって `left` 定義 `top` されます。 これらは、スライドのそれぞれの`left: 0``top: 0`エッジからの余白として機能し、ポイント単位で測定され、左上隅になります。 図形のサイズは、and プロパティで `height` 指定 `width` します。 これらのプロパティをリセットすることで、コードで図形を移動またはサイズ変更できます。 (これらのプロパティは、図形が線の場合、少し異なる意味を持ちます。 「 [Lines](#lines).」を参照してください。

## <a name="text-in-shapes"></a>図形内のテキスト

幾何学的図形には、テキストを含めできます。 図形には、`textFrame`[TextFrame 型のプロパティがあります](/javascript/api/powerpoint/powerpoint.textframe)。 オブジェクト `TextFrame` は、テキスト表示オプション (余白やテキスト オーバーフローなど) を管理します。 `TextFrame.textRange` は、 [テキストコンテンツと](/javascript/api/powerpoint/powerpoint.textrange) フォント設定を持つ TextRange オブジェクトです。

次のコード サンプルでは、" **Shape text" というテキストを持つ "Braces"** という名前の幾何学的 **な図形を作成します**。 また、図形とテキストの色を調整し、テキストの垂直方向の配置を中央に設定します。

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

図形は、オブジェクトのメソッドを使用してスライド `Shape` から削除 `delete` されます。

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
