---
title: Excel JavaScript API を使用して図形を操作する
description: ''
ms.date: 09/03/2019
localization_priority: Normal
ms.openlocfilehash: 2461416bcd7e64c2ea300d98e504ff27edcb14ac
ms.sourcegitcommit: 78998a9f0ebb81c4dd2b77574148b16fe6725cfc
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/03/2019
ms.locfileid: "36715607"
---
# <a name="work-with-shapes-using-the-excel-javascript-api"></a><span data-ttu-id="82cdb-102">Excel JavaScript API を使用して図形を操作する</span><span class="sxs-lookup"><span data-stu-id="82cdb-102">Work with shapes using the Excel JavaScript API</span></span>

<span data-ttu-id="82cdb-103">Excel では、図形は Excel の描画層にある任意のオブジェクトとして定義されます。</span><span class="sxs-lookup"><span data-stu-id="82cdb-103">Excel defines shapes as any object that sits on the drawing layer of Excel.</span></span> <span data-ttu-id="82cdb-104">つまり、セルの外部にあるものは図形です。</span><span class="sxs-lookup"><span data-stu-id="82cdb-104">That means anything outside of a cell is a shape.</span></span> <span data-ttu-id="82cdb-105">この記事では、[図形](/javascript/api/excel/excel.shape)および shapes [ecollection](/javascript/api/excel/excel.shapecollection) api と組み合わせて、ジオメトリック図形、線、およびイメージを使用する方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="82cdb-105">This article describes how to use geometric shapes, lines, and images in conjunction with the [Shape](/javascript/api/excel/excel.shape) and [ShapeCollection](/javascript/api/excel/excel.shapecollection) APIs.</span></span> <span data-ttu-id="82cdb-106">[グラフ](/javascript/api/excel/excel.chart)については、「 [Excel JavaScript API を使用してグラフを操作](excel-add-ins-charts.md)する」で説明されています。</span><span class="sxs-lookup"><span data-stu-id="82cdb-106">[Charts](/javascript/api/excel/excel.chart) are covered in their own article, [Work with charts using the Excel JavaScript API](excel-add-ins-charts.md).</span></span>

## <a name="create-shapes"></a><span data-ttu-id="82cdb-107">図形を作成する</span><span class="sxs-lookup"><span data-stu-id="82cdb-107">Create shapes</span></span>

<span data-ttu-id="82cdb-108">図形は、ワークシートの shape コレクション (`Worksheet.shapes`) を使用して作成され、格納されます。</span><span class="sxs-lookup"><span data-stu-id="82cdb-108">Shapes are created through and stored in a worksheet's shape collection (`Worksheet.shapes`).</span></span> <span data-ttu-id="82cdb-109">`ShapeCollection`には`.add*` 、この目的のためにいくつかの方法があります。</span><span class="sxs-lookup"><span data-stu-id="82cdb-109">`ShapeCollection` has several `.add*` methods for this purpose.</span></span> <span data-ttu-id="82cdb-110">すべての図形には、コレクションに追加されたときに名前と Id が生成されます。</span><span class="sxs-lookup"><span data-stu-id="82cdb-110">All shapes have names and IDs generated for them when they are added to the collection.</span></span> <span data-ttu-id="82cdb-111">これらは、 `name`および`id`プロパティです。</span><span class="sxs-lookup"><span data-stu-id="82cdb-111">These are the `name` and `id` properties, respectively.</span></span> <span data-ttu-id="82cdb-112">`name`アドインで設定して、 `ShapeCollection.getItem(name)`メソッドを使用して簡単に取得することができます。</span><span class="sxs-lookup"><span data-stu-id="82cdb-112">`name` can be set by your add-in for easy retrieval with the `ShapeCollection.getItem(name)` method.</span></span>

<span data-ttu-id="82cdb-113">次の種類の図形は、関連付けられているメソッドを使用して追加されます。</span><span class="sxs-lookup"><span data-stu-id="82cdb-113">The following types of shapes are added using the associated method:</span></span>

| <span data-ttu-id="82cdb-114">Shape</span><span class="sxs-lookup"><span data-stu-id="82cdb-114">Shape</span></span> | <span data-ttu-id="82cdb-115">Tabs.Add メソッド (Outlook フォーム スクリプト)</span><span class="sxs-lookup"><span data-stu-id="82cdb-115">Add Method</span></span> | <span data-ttu-id="82cdb-116">署名</span><span class="sxs-lookup"><span data-stu-id="82cdb-116">Signature</span></span> |
|-------|------------|-----------|
| <span data-ttu-id="82cdb-117">幾何学的図形</span><span class="sxs-lookup"><span data-stu-id="82cdb-117">Geometric Shape</span></span> | [<span data-ttu-id="82cdb-118">addGeometricShape</span><span class="sxs-lookup"><span data-stu-id="82cdb-118">addGeometricShape</span></span>](/javascript/api/excel/excel.shapecollection#addgeometricshape-geometricshapetype-) | `addGeometricShape(geometricShapeType: Excel.GeometricShapeType): Excel.Shape` |
| <span data-ttu-id="82cdb-119">画像 (JPEG または PNG のいずれか)</span><span class="sxs-lookup"><span data-stu-id="82cdb-119">Image (either JPEG or PNG)</span></span> | [<span data-ttu-id="82cdb-120">addImage</span><span class="sxs-lookup"><span data-stu-id="82cdb-120">addImage</span></span>](/javascript/api/excel/excel.shapecollection#addimage-base64imagestring-) | `addImage(base64ImageString: string): Excel.Shape` |
| <span data-ttu-id="82cdb-121">枠線</span><span class="sxs-lookup"><span data-stu-id="82cdb-121">Line</span></span> | [<span data-ttu-id="82cdb-122">addLine</span><span class="sxs-lookup"><span data-stu-id="82cdb-122">addLine</span></span>](/javascript/api/excel/excel.shapecollection#addline-startleft--starttop--endleft--endtop--connectortype-) | `addLine(startLeft: number, startTop: number, endLeft: number, endTop: number, connectorType?: Excel.ConnectorType): Excel.Shape` |
| <span data-ttu-id="82cdb-123">SVG</span><span class="sxs-lookup"><span data-stu-id="82cdb-123">SVG</span></span> | [<span data-ttu-id="82cdb-124">addSvg</span><span class="sxs-lookup"><span data-stu-id="82cdb-124">addSvg</span></span>](/javascript/api/excel/excel.shapecollection#addsvg-xml-) | `addSvg(xml: string): Excel.Shape` |
| <span data-ttu-id="82cdb-125">テキスト ボックス</span><span class="sxs-lookup"><span data-stu-id="82cdb-125">Text Box</span></span> | [<span data-ttu-id="82cdb-126">addTextBox</span><span class="sxs-lookup"><span data-stu-id="82cdb-126">addTextBox</span></span>](/javascript/api/excel/excel.shapecollection#addtextbox-text-) | `addTextBox(text?: string): Excel.Shape` |

### <a name="geometric-shapes"></a><span data-ttu-id="82cdb-127">幾何学的な図形</span><span class="sxs-lookup"><span data-stu-id="82cdb-127">Geometric shapes</span></span>

<span data-ttu-id="82cdb-128">ジオメトリック図形が作成され`ShapeCollection.addGeometricShape`ます。</span><span class="sxs-lookup"><span data-stu-id="82cdb-128">A geometric shape is created with `ShapeCollection.addGeometricShape`.</span></span> <span data-ttu-id="82cdb-129">このメソッドは、 [GeometricShapeType](/javascript/api/excel/excel.geometricshapetype) enum を引数として受け取ります。</span><span class="sxs-lookup"><span data-stu-id="82cdb-129">That method takes a [GeometricShapeType](/javascript/api/excel/excel.geometricshapetype) enum as an argument.</span></span>

<span data-ttu-id="82cdb-130">次のコードサンプルでは、ワークシートの上端と左端から100ピクセルに配置された、 **"Square"** という名前の150x150 ピクセルの四角形を作成します。</span><span class="sxs-lookup"><span data-stu-id="82cdb-130">The following code sample creates a 150x150-pixel rectangle named **"Square"** that is positioned 100 pixels from the top and left sides of the worksheet.</span></span>

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

### <a name="images"></a><span data-ttu-id="82cdb-131">画像</span><span class="sxs-lookup"><span data-stu-id="82cdb-131">Images</span></span>

<span data-ttu-id="82cdb-132">JPEG、PNG、SVG の画像は、図形としてワークシートに挿入できます。</span><span class="sxs-lookup"><span data-stu-id="82cdb-132">JPEG, PNG, and SVG images can be inserted into a worksheet as shapes.</span></span> <span data-ttu-id="82cdb-133">メソッド`ShapeCollection.addImage`は、base64 でエンコードされた文字列を引数として受け取ります。</span><span class="sxs-lookup"><span data-stu-id="82cdb-133">The `ShapeCollection.addImage` method takes a base64-encoded string as an argument.</span></span> <span data-ttu-id="82cdb-134">これは、文字列形式の JPEG または PNG 画像のいずれかです。</span><span class="sxs-lookup"><span data-stu-id="82cdb-134">This is either a JPEG or PNG image in string form.</span></span> <span data-ttu-id="82cdb-135">`ShapeCollection.addSvg`も文字列で受け取りますが、この引数はグラフィックを定義する XML です。</span><span class="sxs-lookup"><span data-stu-id="82cdb-135">`ShapeCollection.addSvg` also takes in a string, though this argument is XML that defines the graphic.</span></span>

<span data-ttu-id="82cdb-136">次のコードサンプルは、 [FileReader](https://developer.mozilla.org/docs/Web/API/FileReader)によって、文字列としてロードされているイメージファイルを示しています。</span><span class="sxs-lookup"><span data-stu-id="82cdb-136">The following code sample shows an image file being loaded by a [FileReader](https://developer.mozilla.org/docs/Web/API/FileReader) as a string.</span></span> <span data-ttu-id="82cdb-137">文字列には、図形が作成される前に、メタデータ "base64" が削除されています。</span><span class="sxs-lookup"><span data-stu-id="82cdb-137">The string has the metadata "base64," removed before the shape is created.</span></span>

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

### <a name="lines"></a><span data-ttu-id="82cdb-138">Lines</span><span class="sxs-lookup"><span data-stu-id="82cdb-138">Lines</span></span>

<span data-ttu-id="82cdb-139">行はに`ShapeCollection.addLine`作成されます。</span><span class="sxs-lookup"><span data-stu-id="82cdb-139">A line is created with `ShapeCollection.addLine`.</span></span> <span data-ttu-id="82cdb-140">このメソッドには、行の開始点と終了点の左余白と上余白が必要です。</span><span class="sxs-lookup"><span data-stu-id="82cdb-140">That method needs the left and top margins of the line's start and end points.</span></span> <span data-ttu-id="82cdb-141">また、 [ConnectorType](/javascript/api/excel/excel.connectortype)列挙を取得して、エンドポイント間の行の contorts 方法を指定します。</span><span class="sxs-lookup"><span data-stu-id="82cdb-141">It also takes a [ConnectorType](/javascript/api/excel/excel.connectortype) enum to specify how the line contorts between endpoints.</span></span> <span data-ttu-id="82cdb-142">次のコードサンプルでは、ワークシートに直線を作成します。</span><span class="sxs-lookup"><span data-stu-id="82cdb-142">The following code sample creates a straight line on the worksheet.</span></span>

```js
// This sample creates a straight line from [200,50] to [300,150] on the worksheet
Excel.run(function (context) {
    var shapes = context.workbook.worksheets.getItem("MyWorksheet").shapes;
    var line = shapes.addLine(200, 50, 300, 150, Excel.ConnectorType.straight);
    line.name = "StraightLine";
    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="82cdb-143">線は、他の Shape オブジェクトに接続することができます。</span><span class="sxs-lookup"><span data-stu-id="82cdb-143">Lines can be connected to other Shape objects.</span></span> <span data-ttu-id="82cdb-144">メソッド`connectBeginShape`と`connectEndShape`メソッドは、指定された接続ポイントにある図形に対して、線の始点と終点を接続します。</span><span class="sxs-lookup"><span data-stu-id="82cdb-144">The `connectBeginShape` and `connectEndShape` methods attach the beginning and ending of a line to shapes at the specified connection points.</span></span> <span data-ttu-id="82cdb-145">これらのポイントの位置は図形によって異なり`Shape.connectionSiteCount`ますが、を使用すると、アドインが範囲外のポイントに接続されないようにすることができます。</span><span class="sxs-lookup"><span data-stu-id="82cdb-145">The locations of these points vary by shape, but the `Shape.connectionSiteCount` can be used to ensure your add-in does not connect to a point that's out-of-bounds.</span></span> <span data-ttu-id="82cdb-146">`disconnectBeginShape`および`disconnectEndShape`メソッドを使用して、接続されているすべての図形から線が切断されます。</span><span class="sxs-lookup"><span data-stu-id="82cdb-146">A line is disconnected from any attached shapes using the `disconnectBeginShape` and `disconnectEndShape` methods.</span></span>

<span data-ttu-id="82cdb-147">次のコードサンプルでは、" **Myline"** 行を **"l shape"** と **"直角図形"** という名前の2つの図形に接続します。</span><span class="sxs-lookup"><span data-stu-id="82cdb-147">The following code sample connects the **"MyLine"** line to two shapes named **"LeftShape"** and **"RightShape"**.</span></span>

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

## <a name="move-and-resize-shapes"></a><span data-ttu-id="82cdb-148">図形の移動とサイズ変更</span><span class="sxs-lookup"><span data-stu-id="82cdb-148">Move and resize shapes</span></span>

<span data-ttu-id="82cdb-149">ワークシートの一番上にある図形。</span><span class="sxs-lookup"><span data-stu-id="82cdb-149">Shapes sit on top of the worksheet.</span></span> <span data-ttu-id="82cdb-150">これらの配置は、 `left`および`top`プロパティによって定義されます。</span><span class="sxs-lookup"><span data-stu-id="82cdb-150">Their placement is defined by the `left` and `top` property.</span></span> <span data-ttu-id="82cdb-151">これらは、ワークシートの各エッジの余白として機能し、[0, 0] が左上隅になります。</span><span class="sxs-lookup"><span data-stu-id="82cdb-151">These act as margins from worksheet's respective edges, with [0, 0] being the upper-left corner.</span></span> <span data-ttu-id="82cdb-152">これらは、 `incrementLeft`および`incrementTop`メソッドを使用して、現在の位置から直接設定または調整することができます。</span><span class="sxs-lookup"><span data-stu-id="82cdb-152">These can either be set directly or adjusted from their current position with the `incrementLeft` and `incrementTop` methods.</span></span> <span data-ttu-id="82cdb-153">既定の位置から図形を回転させる度合いは、この方法で設定します。 `rotation`この方法では、プロパティが`incrementRotation`絶対量で、既存の回転を調整するメソッドも使用されます。</span><span class="sxs-lookup"><span data-stu-id="82cdb-153">How much a shape is rotated from the default position is also established in this manner, with the `rotation` property being the absolute amount and the `incrementRotation` method adjusting the existing rotation.</span></span>

<span data-ttu-id="82cdb-154">他の図形を基準とした図形の深さは`zorderPosition` 、プロパティによって定義されます。</span><span class="sxs-lookup"><span data-stu-id="82cdb-154">A shape's depth relative to other shapes is defined by the `zorderPosition` property.</span></span> <span data-ttu-id="82cdb-155">これは`setZOrder`メソッドを使用して設定され[](/javascript/api/excel/excel.shapezorder)ます。このメソッドは、このメソッドを使用します。</span><span class="sxs-lookup"><span data-stu-id="82cdb-155">This is set using the `setZOrder` method, which takes a [ShapeZOrder](/javascript/api/excel/excel.shapezorder).</span></span> <span data-ttu-id="82cdb-156">`setZOrder`他の図形を基準に現在の図形の順序を調整します。</span><span class="sxs-lookup"><span data-stu-id="82cdb-156">`setZOrder` adjusts the ordering of the current shape relative to the other shapes.</span></span>

<span data-ttu-id="82cdb-157">アドインには、図形の高さと幅を変更するためのいくつかのオプションがあります。</span><span class="sxs-lookup"><span data-stu-id="82cdb-157">Your add-in has a couple options for changing the height and width of shapes.</span></span> <span data-ttu-id="82cdb-158">`height`または`width`プロパティのいずれかを設定すると、他の次元を変更せずに、指定した次元が変更されます。</span><span class="sxs-lookup"><span data-stu-id="82cdb-158">Setting either the `height` or `width` property changes the specified dimension without changing the other dimension.</span></span> <span data-ttu-id="82cdb-159">`scaleHeight`を指定`scaleWidth`して、現在のサイズまたは元のサイズを基準にして図形のそれぞれの寸法を調整します (提供される[ShapeScaleType](/javascript/api/excel/excel.shapescaletype)の値に基づきます)。</span><span class="sxs-lookup"><span data-stu-id="82cdb-159">The `scaleHeight` and `scaleWidth` adjust the shape's respective dimensions relative to either the current or original size (based on the value of the provided [ShapeScaleType](/javascript/api/excel/excel.shapescaletype)).</span></span> <span data-ttu-id="82cdb-160">省略可能な[ShapeScaleFrom](/javascript/api/excel/excel.shapescalefrom)パラメーターは、図形を拡大または縮小する位置 (左上隅、中央、または右下隅) を指定します。</span><span class="sxs-lookup"><span data-stu-id="82cdb-160">An optional [ShapeScaleFrom](/javascript/api/excel/excel.shapescalefrom) parameter specifies from where the shape scales (top-left corner, middle, or bottom-right corner).</span></span> <span data-ttu-id="82cdb-161">プロパティに`lockAspectRatio` **true**が設定されている場合、scale メソッドは、他の次元も調整して、図形の現在の縦横比を維持します。</span><span class="sxs-lookup"><span data-stu-id="82cdb-161">If the `lockAspectRatio` property is **true**, the scale methods maintain the shape's current aspect ratio by also adjusting the other dimension.</span></span>

> [!NOTE]
> <span data-ttu-id="82cdb-162">プロパティに`height` `width`対する直接の変更は、プロパティの値に関係なく`lockAspectRatio` 、そのプロパティにのみ影響します。</span><span class="sxs-lookup"><span data-stu-id="82cdb-162">Direct changes to the `height` and `width` properties only affect that property, regardless of the `lockAspectRatio` property's value.</span></span>

<span data-ttu-id="82cdb-163">次のコードサンプルでは、元のサイズに1.25 倍に拡大または縮小された図形を表示します。</span><span class="sxs-lookup"><span data-stu-id="82cdb-163">The following code sample shows a shape being scaled to 1.25 times its original size and rotated 30 degrees.</span></span>

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

## <a name="text-in-shapes"></a><span data-ttu-id="82cdb-164">図形内のテキスト</span><span class="sxs-lookup"><span data-stu-id="82cdb-164">Text in shapes</span></span>

<span data-ttu-id="82cdb-165">幾何学的な図形にはテキストを含めることができます。</span><span class="sxs-lookup"><span data-stu-id="82cdb-165">Geometric Shapes can contain text.</span></span> <span data-ttu-id="82cdb-166">図形には`textFrame` 、 [TextFrame](/javascript/api/excel/excel.textframe)型のプロパティがあります。</span><span class="sxs-lookup"><span data-stu-id="82cdb-166">Shapes have a `textFrame` property of type [TextFrame](/javascript/api/excel/excel.textframe).</span></span> <span data-ttu-id="82cdb-167">オブジェクト`TextFrame`は、テキスト表示オプション (余白、テキストオーバーフローなど) を管理します。</span><span class="sxs-lookup"><span data-stu-id="82cdb-167">The `TextFrame` object manages the text display options (such as margins and text overflow).</span></span> <span data-ttu-id="82cdb-168">`TextFrame.textRange`は、テキストの内容とフォントの設定を含む[TextRange](/javascript/api/excel/excel.textrange)オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="82cdb-168">`TextFrame.textRange` is a [TextRange](/javascript/api/excel/excel.textrange) object with the text content and font settings.</span></span>

<span data-ttu-id="82cdb-169">次のコードサンプルでは、テキスト "Shape Text" を使用して "Wave" という名前のジオメトリック図形を作成します。</span><span class="sxs-lookup"><span data-stu-id="82cdb-169">The following code sample creates a geometric shape named "Wave" with the text "Shape Text".</span></span> <span data-ttu-id="82cdb-170">また、図形とテキストの色を調整するだけでなく、テキストの水平方向の配置を中央に設定します。</span><span class="sxs-lookup"><span data-stu-id="82cdb-170">It also adjusts the shape and text colors, as well as sets the text's horizontal alignment to the center.</span></span>

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

<span data-ttu-id="82cdb-171">の`addTextBox` `ShapeCollection`メソッドは、白`GeometricShape`の背景`Rectangle`と黒のテキストを使用して、型を作成します。</span><span class="sxs-lookup"><span data-stu-id="82cdb-171">The `addTextBox` method of `ShapeCollection` creates a `GeometricShape` of type `Rectangle` with a white background and black text.</span></span> <span data-ttu-id="82cdb-172">これは、[挿入] タブの [Excel の**テキストボックス**によって\*\*\*\* 作成され`addTextBox`たものと同じです。文字列型 (string `TextRange`) の引数を指定して、のテキストを設定します。</span><span class="sxs-lookup"><span data-stu-id="82cdb-172">This is the same as what is created by Excel's **Text Box** button on the **Insert** tab. `addTextBox` takes a string argument to set the text of the `TextRange`.</span></span>

<span data-ttu-id="82cdb-173">次のコードサンプルは、"Hello!" というテキストを含むテキストボックスを作成する方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="82cdb-173">The following code sample shows the creation of a text box with the text "Hello!".</span></span>

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

## <a name="shape-groups"></a><span data-ttu-id="82cdb-174">図形グループ</span><span class="sxs-lookup"><span data-stu-id="82cdb-174">Shape groups</span></span>

<span data-ttu-id="82cdb-175">図形は一緒にグループ化できます。</span><span class="sxs-lookup"><span data-stu-id="82cdb-175">Shapes can be grouped together.</span></span> <span data-ttu-id="82cdb-176">これにより、ユーザーは、配置、サイズ変更、およびその他の関連タスクのために1つのエンティティとして扱うことができます。</span><span class="sxs-lookup"><span data-stu-id="82cdb-176">This allows a user to treat them as a single entity for positioning, sizing, and other related tasks.</span></span> <span data-ttu-id="82cdb-177">図形[グループ](/javascript/api/excel/excel.shapegroup)はの`Shape`種類であるため、アドインでグループを1つの図形として扱うことができます。</span><span class="sxs-lookup"><span data-stu-id="82cdb-177">A [ShapeGroup](/javascript/api/excel/excel.shapegroup) is a type of `Shape`, so your add-in treats the group as a single shape.</span></span>

<span data-ttu-id="82cdb-178">次のコードサンプルでは、グループ化された3つの図形を示します。</span><span class="sxs-lookup"><span data-stu-id="82cdb-178">The following code sample shows three shapes being grouped together.</span></span> <span data-ttu-id="82cdb-179">次のコードサンプルでは、図形グループを50ピクセル右に移動していることを示します。</span><span class="sxs-lookup"><span data-stu-id="82cdb-179">The subsequent code sample shows that shape group being moved to the right 50 pixels.</span></span>

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
> <span data-ttu-id="82cdb-180">グループ内の個々の図形は、種類`ShapeGroup.shapes`が[groupshapecollection](/javascript/api/excel/excel.GroupShapeCollection)であるプロパティを介して参照されます。</span><span class="sxs-lookup"><span data-stu-id="82cdb-180">Individual shapes within the group are referenced through the `ShapeGroup.shapes` property, which is of type [GroupShapeCollection](/javascript/api/excel/excel.GroupShapeCollection).</span></span> <span data-ttu-id="82cdb-181">グループ化された後は、ワークシートの shape コレクションからアクセスできなくなります。</span><span class="sxs-lookup"><span data-stu-id="82cdb-181">They are no longer accessible through the worksheet's shape collection after being grouped.</span></span> <span data-ttu-id="82cdb-182">たとえば、ワークシートに3つの図形があり、すべてが一緒にグループ化され`shapes.getCount`ている場合、ワークシートのメソッドはカウントを1とします。</span><span class="sxs-lookup"><span data-stu-id="82cdb-182">As an example, if your worksheet had three shapes and they were all grouped together, the worksheet's `shapes.getCount` method would return a count of 1.</span></span>

## <a name="export-shapes-as-images"></a><span data-ttu-id="82cdb-183">図形を画像としてエクスポートする</span><span class="sxs-lookup"><span data-stu-id="82cdb-183">Export shapes as images</span></span>

<span data-ttu-id="82cdb-184">任意`Shape`のオブジェクトをイメージに変換できます。</span><span class="sxs-lookup"><span data-stu-id="82cdb-184">Any `Shape` object can be converted to an image.</span></span> <span data-ttu-id="82cdb-185">[GetAsImage](/javascript/api/excel/excel.shape#getasimage-format-)は、base64 でエンコードされた文字列を返します。</span><span class="sxs-lookup"><span data-stu-id="82cdb-185">[Shape.getAsImage](/javascript/api/excel/excel.shape#getasimage-format-) returns base64-encoded string.</span></span> <span data-ttu-id="82cdb-186">画像の形式は、に`getAsImage`渡される図[形式](/javascript/api/excel/excel.pictureformat)の列挙体として指定されます。</span><span class="sxs-lookup"><span data-stu-id="82cdb-186">The image's format is specified as a [PictureFormat](/javascript/api/excel/excel.pictureformat) enum passed to `getAsImage`.</span></span>

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

## <a name="delete-shapes"></a><span data-ttu-id="82cdb-187">図形を削除する</span><span class="sxs-lookup"><span data-stu-id="82cdb-187">Delete shapes</span></span>

<span data-ttu-id="82cdb-188">図形は、 `Shape`オブジェクトの`delete`メソッドを使用してワークシートから削除されます。</span><span class="sxs-lookup"><span data-stu-id="82cdb-188">Shapes are removed from the worksheet with the `Shape` object's `delete` method.</span></span> <span data-ttu-id="82cdb-189">その他のメタデータは必要ありません。</span><span class="sxs-lookup"><span data-stu-id="82cdb-189">No other metadata is needed.</span></span>

<span data-ttu-id="82cdb-190">次のコードサンプルでは、 **Myworksheet**からすべての図形を削除します。</span><span class="sxs-lookup"><span data-stu-id="82cdb-190">The following code sample deletes all the shapes from **MyWorksheet**.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="82cdb-191">関連項目</span><span class="sxs-lookup"><span data-stu-id="82cdb-191">See also</span></span>

- [<span data-ttu-id="82cdb-192">Excel JavaScript API を使用した基本的なプログラミングの概念</span><span class="sxs-lookup"><span data-stu-id="82cdb-192">Fundamental programming concepts with the Excel JavaScript API</span></span>](../reference/overview/excel-add-ins-reference-overview.md)
- [<span data-ttu-id="82cdb-193">Excel JavaScript API を使用してグラフを操作する</span><span class="sxs-lookup"><span data-stu-id="82cdb-193">Work with charts using the Excel JavaScript API</span></span>](excel-add-ins-charts.md)
