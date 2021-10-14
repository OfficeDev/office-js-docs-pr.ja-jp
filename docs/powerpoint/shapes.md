---
title: JavaScript API を使用して図形PowerPointする
description: スライドに図形を追加、削除、書式設定するPowerPointします。
ms.date: 10/06/2021
ms.localizationpriority: medium
ms.openlocfilehash: e510ff47f4c54cd465be5c97c5828aad81041c5e
ms.sourcegitcommit: fb4a55764fb60e826ad06d15d1539e41df503b65
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/14/2021
ms.locfileid: "60356864"
---
# <a name="work-with-shapes-using-the-powerpoint-javascript-api-preview"></a>JavaScript API を使用して図形をPowerPointする (プレビュー)

この記事では、図形および[ShapeCollection](/javascript/api/powerpoint/powerpoint.shapecollection) API と組み合わせて、[](/javascript/api/powerpoint/powerpoint.shape)幾何学的図形、線、およびテキスト ボックスを使用する方法について説明します。

[!INCLUDE [Information about using preview APIs](../includes/using-preview-apis-host.md)]

## <a name="create-shapes"></a>図形を作成する

図形は、スライドの図形コレクション () を通じて作成され、格納されます `slide.shapes` 。 `ShapeCollection` この目的 `.add*` のためにいくつかのメソッドがあります。 すべての図形には、コレクションに追加するときに名前と ID が生成されます。 これらは、それぞれ `name` プロパティ `id` とプロパティです。 `name` をアドインで設定できます。

### <a name="geometric-shapes"></a>幾何学的図形

幾何学的な図形は、 のオーバーロードの 1 つで作成されます `ShapeCollection.addGeometricShape` 。 最初のパラメーターは [、GeometricShapeType](/javascript/api/powerpoint/powerpoint.geometricshapetype) 列挙型か、列挙型の値のいずれかと同等の文字列です。 [ShapeAddOptions](/javascript/api/powerpoint/powerpoint.shapeaddoptions)型のオプションの 2 番目のパラメーターを使用して、スライドの上辺と左側を基準に図形の初期サイズとその位置をポイント単位で指定できます。 または、図形の作成後にこれらのプロパティを設定できます。

次のコード サンプルでは、スライドの上辺と左側から 100 ポイントの位置にある **"Square"** という名前の四角形を作成します。 メソッドはオブジェクトを返 `Shape` します。

```js
// This sample creates a rectangle positioned 100 points from the top and left sides
// of the slide and is 150x150 points. The shape is put on the first slide.
PowerPoint.run(function (context) {
    var shapes = context.presentation.slides.getItemAt(0).shapes;
    var rectangle = shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle);
    rectangle.left = 100;
    rectangle.top = 100;
    rectangle.height = 150;
    rectangle.width = 150;
    rectangle.name = "Square";
    return context.sync();
});
```

### <a name="lines"></a>Lines

行は、 のオーバーロードの 1 つで作成されます `ShapeCollection.addLine` 。 最初のパラメーターは [、ConnectorType](/javascript/api/powerpoint/powerpoint.connectortype) 列挙型か、列挙型の値のいずれかと同等の文字列で、エンドポイント間の行のコントルト方法を指定します。 線の開始点と終了点を指定できる [、ShapeAddOptions](/javascript/api/powerpoint/powerpoint.shapeaddoptions) 型のオプションの 2 番目のパラメーターがあります。 または、図形の作成後にこれらのプロパティを設定できます。 メソッドはオブジェクトを返 `Shape` します。

> [!NOTE]
> 図形が線の場合、and オブジェクトのプロパティは、スライドの上端と左端を基準に線の開始点 `top` `left` `Shape` `ShapeAddOptions` を指定します。 and `height` プロパティ `width` は、開始点を基準に線 *の終点を指定します*。 したがって、スライドの上端と左端を基準にした終了点は 、 ( `top`  +  `height` ) です `left`  +  `width` 。 すべてのプロパティの測定単位はポイントであり、負の値を使用できます。

次のコード サンプルでは、スライドに直線を作成します。

```js
// This sample creates a straight line on the first slide.
PowerPoint.run(function (context) {
    var shapes = context.presentation.slides.getItemAt(0).shapes;
    var line = shapes.addLine(Excel.ConnectorType.straight, {left: 200, top: 50, height: 300, width: 150});
    line.name = "StraightLine";
    return context.sync();
});
```

### <a name="text-boxes"></a>テキスト ボックス

addTextBox メソッドを使用して [テキスト ボックスが作成](/javascript/api/powerpoint/powerpoint.shapecollection#addTextBox_text__options_) されます。 最初のパラメーターは、最初にボックスに表示されるテキストです。 [ShapeAddOptions](/javascript/api/powerpoint/powerpoint.shapeaddoptions)型のオプションの 2 番目のパラメーターを使用して、テキスト ボックスの初期サイズと、スライドの上辺と左側の位置を指定できます。 または、図形の作成後にこれらのプロパティを設定できます。

次のコード サンプルは、最初のスライドにテキスト ボックスを作成する方法を示しています。

```js
// This sample creates a text box with the text "Hello!" and sizes it appropriately.
PowerPoint.run(function (context) {
    var shapes = context.presentation.slides.getItemAt(0).shapes;
    var textbox = shapes.addTextBox("Hello!");
    textbox.left = 100;
    textbox.top = 100;
    textbox.height = 300;
    textbox.width = 450;
    textbox.name = "Textbox";
    return context.sync();
});
```

## <a name="move-and-resize-shapes"></a>図形の移動とサイズ変更

図形はスライドの上に座っています。 それらの配置は、and プロパティ `left` によって定義 `top` されます。 これらは、スライドのそれぞれのエッジからの余白として機能し、ポイント単位で測定され、 `left: 0` 左上 `top: 0` 隅になります。 図形のサイズは、and プロパティで `height` 指定 `width` します。 これらのプロパティをリセットすることで、コードで図形を移動またはサイズ変更できます。 (これらのプロパティは、図形が線の場合、少し異なる意味を持ちます。 「Lines [.)」を参照](#lines)してください。

## <a name="text-in-shapes"></a>図形内のテキスト

幾何学的図形には、テキストを含めできます。 図形には `textFrame` 、TextFrame 型の [プロパティがあります](/javascript/api/powerpoint/powerpoint.textframe)。 オブジェクト `TextFrame` は、テキスト表示オプション (余白やテキスト オーバーフローなど) を管理します。 `TextFrame.textRange` は、 [テキストコンテンツと](/javascript/api/powerpoint/powerpoint.textrange) フォント設定を持つ TextRange オブジェクトです。

次のコード サンプルでは、"Shape **text" というテキストを持つ "Braces"** という名前の幾何学的 **図形を作成します**。 また、図形とテキストの色を調整し、テキストの垂直方向の配置を中央に設定します。

```js
// This sample creates a light blue rectangle with braces ("{}") on the left and right ends and adds the purple text "Shape text" to the center.
PowerPoint.run(function (context) {
    var shapes = context.presentation.slides.getItemAt(0).shapes;
    var braces = shapes.addGeometricShape(PowerPoint.GeometricShapeType.bracePair);
    braces.left = 100;
    braces.top = 400;
    braces.height = 50;
    braces.width = 150;
    braces.name = "Braces";
    braces.fill.setSolidColor("lightblue");
    braces.textFrame.textRange.text = "Shape text";
    braces.textFrame.textRange.font.color = "purple";
    braces.textFrame.verticalAlignment = PowerPoint.TextVerticalAlignment.middleCentered;
    return context.sync();
});
```

## <a name="delete-shapes"></a>図形を削除する

図形は、オブジェクトのメソッドを使用して `Shape` スライドから削除 `delete` されます。

次のコード サンプルは、図形を削除する方法を示しています。

```js
PowerPoint.run(function (context) {
    // Delete all shapes from the first slide.
    var sheet = context.presentation.slides.getItemAt(0);
    var shapes = sheet.shapes;

    // Load all the shapes in the collection without loading their properties.
    shapes.load("items/$none");
    return context.sync()
        .then(function () {
            shapes.items.forEach(function (shape) {
                shape.delete()
            });
            return context.sync();
        })
       .catch(errorHandlerFunction);
});
```
