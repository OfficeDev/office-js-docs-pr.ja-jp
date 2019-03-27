---
title: Excel JavaScript API を使用して図形を操作する
description: ''
ms.date: 03/21/2019
localization_priority: Normal
ms.openlocfilehash: e4d01c387fff01d68cb26369240a1e06e723a54c
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/27/2019
ms.locfileid: "30926671"
---
# <a name="work-with-shapes-using-the-excel-javascript-api-preview"></a><span data-ttu-id="a31bd-102">Excel JavaScript API を使用して図形を操作する (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="a31bd-102">Work with Shapes using the Excel JavaScript API (preview)</span></span>

> [!NOTE]
> <span data-ttu-id="a31bd-103">この記事に記載されている api は、現時点ではパブリックプレビューでのみ利用可能です。</span><span class="sxs-lookup"><span data-stu-id="a31bd-103">The APIs discussed in this article are currently available only in public preview.</span></span> [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]

<span data-ttu-id="a31bd-104">excel では、図形は excel の描画層にある任意のオブジェクトとして定義されます。</span><span class="sxs-lookup"><span data-stu-id="a31bd-104">Excel defines shapes as any object that sits on the drawing layer of Excel.</span></span> <span data-ttu-id="a31bd-105">つまり、セルの外部にあるものは図形です。</span><span class="sxs-lookup"><span data-stu-id="a31bd-105">That means anything outside of a cell is a shape.</span></span> <span data-ttu-id="a31bd-106">この記事では、ジオメトリック図形、線、およびイメージを [Shape]/javascript/api/excel/excel.shape) および [shapes [ecollection](/javascript/api/excel/excel.shapecollection) ] api と組み合わせて使用する方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="a31bd-106">This article describes how to use geometric shapes, lines, and images in conjunction with the [Shape]/javascript/api/excel/excel.shape) and [ShapeCollection](/javascript/api/excel/excel.shapecollection) APIs.</span></span> <span data-ttu-id="a31bd-107">[グラフ](/javascript/api/excel/excel.chart)については、それぞれの記事「excel JavaScript API を使用してグラフを処理する」 (charts.md) を使用してください。</span><span class="sxs-lookup"><span data-stu-id="a31bd-107">[Charts](/javascript/api/excel/excel.chart) are covered in their own article, [Work with Charts using the Excel JavaScript API]](excel-add-ins-charts.md)).</span></span>

## <a name="create-shapes"></a><span data-ttu-id="a31bd-108">図形を作成する</span><span class="sxs-lookup"><span data-stu-id="a31bd-108">Create shapes</span></span>

<span data-ttu-id="a31bd-109">図形は、ワークシートの shape コレクション (`Worksheet.shapes`) を使用して作成され、格納されます。</span><span class="sxs-lookup"><span data-stu-id="a31bd-109">Shapes are created through and stored in a worksheet's shape collection (`Worksheet.shapes`).</span></span> <span data-ttu-id="a31bd-110">`ShapeCollection`には`.add*` 、この目的のためにいくつかの方法があります。</span><span class="sxs-lookup"><span data-stu-id="a31bd-110">`ShapeCollection` has several `.add*` methods for this purpose.</span></span> <span data-ttu-id="a31bd-111">すべての図形には、コレクションに追加されたときに名前と id が生成されます。</span><span class="sxs-lookup"><span data-stu-id="a31bd-111">All shapes have names and IDs generated for them when they are added to the collection.</span></span> <span data-ttu-id="a31bd-112">これらは、 `name`および`id`プロパティです。</span><span class="sxs-lookup"><span data-stu-id="a31bd-112">These are the `name` and `id` properties, respectively.</span></span> <span data-ttu-id="a31bd-113">`name`アドインで設定して、 `ShapeCollection.getItem(name)`メソッドを使用して簡単に取得することができます。</span><span class="sxs-lookup"><span data-stu-id="a31bd-113">`name` can be set by your add-in for easy retrieval with the `ShapeCollection.getItem(name)` method.</span></span>

<span data-ttu-id="a31bd-114">次の種類の図形は、関連付けられているメソッドを使用して追加されます。</span><span class="sxs-lookup"><span data-stu-id="a31bd-114">The following types of shapes are added using the associated method:</span></span>

| <span data-ttu-id="a31bd-115">Shape</span><span class="sxs-lookup"><span data-stu-id="a31bd-115">Shape</span></span> | <span data-ttu-id="a31bd-116">Tabs.Add メソッド (Outlook フォーム スクリプト)</span><span class="sxs-lookup"><span data-stu-id="a31bd-116">Add Method</span></span> | <span data-ttu-id="a31bd-117">署名</span><span class="sxs-lookup"><span data-stu-id="a31bd-117">Signature</span></span> |
|-------|------------|-----------|
| <span data-ttu-id="a31bd-118">幾何学的図形</span><span class="sxs-lookup"><span data-stu-id="a31bd-118">Geometric Shape</span></span> | [<span data-ttu-id="a31bd-119">addgeometricshape</span><span class="sxs-lookup"><span data-stu-id="a31bd-119">addGeometricShape</span></span>](/javascript/api/excel/excel.shapecollection#addgeometricshape-geometricshapetype-) | `addGeometricShape(geometricShapeType: Excel.GeometricShapeType): Excel.Shape` |
| <span data-ttu-id="a31bd-120">画像 (JPEG または PNG のいずれか)</span><span class="sxs-lookup"><span data-stu-id="a31bd-120">Image (either JPEG or PNG)</span></span> | [<span data-ttu-id="a31bd-121">addimage</span><span class="sxs-lookup"><span data-stu-id="a31bd-121">addImage</span></span>](/javascript/api/excel/excel.shapecollection#addimage-base64imagestring-) | `addImage(base64ImageString: string): Excel.Shape` |
| <span data-ttu-id="a31bd-122">枠線</span><span class="sxs-lookup"><span data-stu-id="a31bd-122">Line</span></span> | [<span data-ttu-id="a31bd-123">addLine</span><span class="sxs-lookup"><span data-stu-id="a31bd-123">addLine</span></span>](/javascript/api/excel/excel.shapecollection#addline-startleft--starttop--endleft--endtop--connectortype-) | `addLine(startLeft: number, startTop: number, endLeft: number, endTop: number, connectorType?: Excel.ConnectorType): Excel.Shape` |
| <span data-ttu-id="a31bd-124">SVG</span><span class="sxs-lookup"><span data-stu-id="a31bd-124">SVG</span></span> | [<span data-ttu-id="a31bd-125">addsvg</span><span class="sxs-lookup"><span data-stu-id="a31bd-125">addSvg</span></span>](/javascript/api/excel/excel.shapecollection#addsvg-xml-) | `addSvg(xml: string): Excel.Shape` |
| <span data-ttu-id="a31bd-126">テキスト ボックス</span><span class="sxs-lookup"><span data-stu-id="a31bd-126">Text Box</span></span> | [<span data-ttu-id="a31bd-127">addTextBox</span><span class="sxs-lookup"><span data-stu-id="a31bd-127">addTextBox</span></span>](/javascript/api/excel/excel.shapecollection#addtextbox-text-) | `addTextBox(text?: string): Excel.Shape` |

### <a name="geometric-shapes"></a><span data-ttu-id="a31bd-128">幾何学的な図形</span><span class="sxs-lookup"><span data-stu-id="a31bd-128">Geometric shapes</span></span>

<span data-ttu-id="a31bd-129">ジオメトリック図形が作成され`ShapeCollection.addGeometricShape`ます。</span><span class="sxs-lookup"><span data-stu-id="a31bd-129">A geometric shape is created with `ShapeCollection.addGeometricShape`.</span></span> <span data-ttu-id="a31bd-130">このメソッドは、 [GeometricShapeType](/javascript/api/excel/excel.geometricshapetype) enum を引数として受け取ります。</span><span class="sxs-lookup"><span data-stu-id="a31bd-130">That method takes a [GeometricShapeType](/javascript/api/excel/excel.geometricshapetype) enum as an argument.</span></span>

<span data-ttu-id="a31bd-131">次のコードサンプルでは、ワークシートの上端と左端から100ピクセルに配置された、 **"Square"** という名前の150x150 ピクセルの四角形を作成します。</span><span class="sxs-lookup"><span data-stu-id="a31bd-131">The following code sample creates a 150x150-pixel rectangle named **"Square"** that is positioned 100 pixels from the top and left sides of the worksheet.</span></span>

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

### <a name="images"></a><span data-ttu-id="a31bd-132">画像</span><span class="sxs-lookup"><span data-stu-id="a31bd-132">Images</span></span>

<span data-ttu-id="a31bd-133">JPEG、PNG、SVG の画像は、図形としてワークシートに挿入できます。</span><span class="sxs-lookup"><span data-stu-id="a31bd-133">JPEG, PNG, and SVG images can be inserted into a worksheet as shapes.</span></span> <span data-ttu-id="a31bd-134">メソッド`ShapeCollection.addImage`は、base64 でエンコードされた文字列を引数として受け取ります。</span><span class="sxs-lookup"><span data-stu-id="a31bd-134">The `ShapeCollection.addImage` method takes a base64-encoded string as an argument.</span></span> <span data-ttu-id="a31bd-135">これは、文字列形式の JPEG または PNG 画像のいずれかです。</span><span class="sxs-lookup"><span data-stu-id="a31bd-135">This is either a JPEG or PNG image in string form.</span></span> <span data-ttu-id="a31bd-136">`ShapeCollection.addSvg`も文字列で受け取りますが、この引数はグラフィックを定義する XML です。</span><span class="sxs-lookup"><span data-stu-id="a31bd-136">`ShapeCollection.addSvg` also takes in a string, though this argument is XML that defines the graphic.</span></span>

<span data-ttu-id="a31bd-137">次のコードサンプルは、 [FileReader](https://developer.mozilla.org/docs/Web/API/FileReader)によって、文字列としてロードされているイメージファイルを示しています。</span><span class="sxs-lookup"><span data-stu-id="a31bd-137">The following code sample shows an image file being loaded by a [FileReader](https://developer.mozilla.org/docs/Web/API/FileReader) as a string.</span></span> <span data-ttu-id="a31bd-138">文字列には、図形が作成される前に、メタデータ "base64" が削除されています。</span><span class="sxs-lookup"><span data-stu-id="a31bd-138">The string has the metadata "base64," removed before the shape is created.</span></span>

```js
// This sample creates an image as a Shape object in the worksheet.
var myFile = document.getElementById("selectedFile");
var reader = new FileReader();

reader.onload = (event) => {
    Excel.run(function (context) {
        var startIndex = event.target.result.indexOf("base64,");
        var myBase64 = event.target.result.substr(startIndex + 7);
        var sheet = context.workbook.worksheets.getItem("MyWorksheet");
        var image = sheet.shapes.addImage(myBase64);
        image.name = "Image";
        return context.sync();
    }).catch(errorHandlerFunction);
};

// Read in the image file as a data URL.
reader.readAsDataURL(myFile.files[0]);
```

### <a name="lines"></a><span data-ttu-id="a31bd-139">Lines</span><span class="sxs-lookup"><span data-stu-id="a31bd-139">Lines</span></span>

<span data-ttu-id="a31bd-140">行はに`ShapeCollection.addLine`作成されます。</span><span class="sxs-lookup"><span data-stu-id="a31bd-140">A line is created with `ShapeCollection.addLine`.</span></span> <span data-ttu-id="a31bd-141">このメソッドには、行の開始点と終了点の左余白と上余白が必要です。</span><span class="sxs-lookup"><span data-stu-id="a31bd-141">That method needs the left and top margins of the line's start and end points.</span></span> <span data-ttu-id="a31bd-142">また、 [ConnectorType](/javascript/api/excel/excel.connectortype)列挙を取得して、エンドポイント間の行の contorts 方法を指定します。</span><span class="sxs-lookup"><span data-stu-id="a31bd-142">It also takes a [ConnectorType](/javascript/api/excel/excel.connectortype) enum to specify how the line contorts between endpoints.</span></span> <span data-ttu-id="a31bd-143">次のコードサンプルでは、ワークシートに直線を作成します。</span><span class="sxs-lookup"><span data-stu-id="a31bd-143">The following code sample creates a straight line on the worksheet.</span></span>

```js
// This sample creates a straight line from [200,50] to [300,150] on the worksheet
Excel.run(function (context) {
    var shapes = context.workbook.worksheets.getItem("MyWorksheet").shapes;
    var line = shapes.addLine(200, 50, 300, 150, Excel.ConnectorType.straight);
    line.name = "StraightLine";
    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="a31bd-144">線は、他の Shape オブジェクトに接続することができます。</span><span class="sxs-lookup"><span data-stu-id="a31bd-144">Lines can be connected to other Shape objects.</span></span> <span data-ttu-id="a31bd-145">メソッド`connectBeginShape`と`connectEndShape`メソッドは、指定された接続ポイントにある図形に対して、線の始点と終点を接続します。</span><span class="sxs-lookup"><span data-stu-id="a31bd-145">The `connectBeginShape` and `connectEndShape` methods attach the beginning and ending of a line to shapes at the specified connection points.</span></span> <span data-ttu-id="a31bd-146">これらのポイントの位置は図形によって異なり`Shape.connectionSiteCount`ますが、を使用すると、アドインが範囲外のポイントに接続されないようにすることができます。</span><span class="sxs-lookup"><span data-stu-id="a31bd-146">The locations of these points vary by shape, but the `Shape.connectionSiteCount` can be used to ensure your add-in does not connect to a point that's out-of-bounds.</span></span> <span data-ttu-id="a31bd-147">`disconnectBeginShape`および`disconnectEndShape`メソッドを使用して、接続されているすべての図形から線が切断されます。</span><span class="sxs-lookup"><span data-stu-id="a31bd-147">A line is disconnected from any attached shapes using the `disconnectBeginShape` and `disconnectEndShape` methods.</span></span>

<span data-ttu-id="a31bd-148">次のコードサンプルでは、" **myline"** 行を **"l shape"** と **"直角図形"** という名前の2つの図形に接続します。</span><span class="sxs-lookup"><span data-stu-id="a31bd-148">The following code sample connects the **"MyLine"** line to two shapes named **"LeftShape"** and **"RightShape"**.</span></span>

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

## <a name="move-and-resize-shapes"></a><span data-ttu-id="a31bd-149">図形の移動とサイズ変更</span><span class="sxs-lookup"><span data-stu-id="a31bd-149">Move and resize shapes</span></span>

<span data-ttu-id="a31bd-150">ワークシートの一番上にある図形。</span><span class="sxs-lookup"><span data-stu-id="a31bd-150">Shapes sit on top of the worksheet.</span></span> <span data-ttu-id="a31bd-151">これらの配置は、 `left`および`top`プロパティによって定義されます。</span><span class="sxs-lookup"><span data-stu-id="a31bd-151">Their placement is defined by the `left` and `top` property.</span></span> <span data-ttu-id="a31bd-152">これらは、ワークシートの各エッジの余白として機能し、[0, 0] が左上隅になります。</span><span class="sxs-lookup"><span data-stu-id="a31bd-152">These act as margins from worksheet's respective edges, with [0, 0] being the upper-left corner.</span></span> <span data-ttu-id="a31bd-153">これらは、 `incrementLeft`および`incrementTop`メソッドを使用して、現在の位置から直接設定または調整することができます。</span><span class="sxs-lookup"><span data-stu-id="a31bd-153">These can either be set directly or adjusted from their current position with the `incrementLeft` and `incrementTop` methods.</span></span> <span data-ttu-id="a31bd-154">既定の位置から図形を回転させる度合いは、この方法で設定します。 `rotation`この方法では、プロパティが`incrementRotation`絶対量で、既存の回転を調整するメソッドも使用されます。</span><span class="sxs-lookup"><span data-stu-id="a31bd-154">How much a shape is rotated from the default position is also established in this manner, with the `rotation` property being the absolute amount and the `incrementRotation` method adjusting the existing rotation.</span></span>

<span data-ttu-id="a31bd-155">他の図形を基準とした図形の深さは`zorderPosition` 、プロパティによって定義されます。</span><span class="sxs-lookup"><span data-stu-id="a31bd-155">A shape's depth relative to other shapes is defined by the `zorderPosition` property.</span></span> <span data-ttu-id="a31bd-156">これは`setZOrder`メソッドを使用して設定され[](/javascript/api/excel/excel.shapezorder)ます。このメソッドは、このメソッドを使用します。</span><span class="sxs-lookup"><span data-stu-id="a31bd-156">This is set using the `setZOrder` method, which takes a [ShapeZOrder](/javascript/api/excel/excel.shapezorder).</span></span> <span data-ttu-id="a31bd-157">`setZOrder`他の図形を基準に現在の図形の順序を調整します。</span><span class="sxs-lookup"><span data-stu-id="a31bd-157">`setZOrder` adjusts the ordering of the current shape relative to the other shapes.</span></span>

<span data-ttu-id="a31bd-158">アドインには、図形の高さと幅を変更するためのいくつかのオプションがあります。</span><span class="sxs-lookup"><span data-stu-id="a31bd-158">Your add-in has a couple options for changing the height and width of shapes.</span></span> <span data-ttu-id="a31bd-159">`height`または`width`プロパティのいずれかを設定すると、他の次元を変更せずに、指定した次元が変更されます。</span><span class="sxs-lookup"><span data-stu-id="a31bd-159">Setting either the `height` or `width` property changes the specified dimension without changing the other dimension.</span></span> <span data-ttu-id="a31bd-160">`scaleHeight`を指定`scaleWidth`して、現在のサイズまたは元のサイズを基準にして図形のそれぞれの寸法を調整します (提供される[ShapeScaleType](/javascript/api/excel/excel.shapescaletype)の値に基づきます)。</span><span class="sxs-lookup"><span data-stu-id="a31bd-160">The `scaleHeight` and `scaleWidth` adjust the shape's respective dimensions relative to either the current or original size (based on the value of the provided [ShapeScaleType](/javascript/api/excel/excel.shapescaletype)).</span></span> <span data-ttu-id="a31bd-161">省略可能な[ShapeScaleFrom](/javascript/api/excel/excel.shapescalefrom)パラメーターは、図形を拡大または縮小する位置 (左上隅、中央、または右下隅) を指定します。</span><span class="sxs-lookup"><span data-stu-id="a31bd-161">An optional [ShapeScaleFrom](/javascript/api/excel/excel.shapescalefrom) parameter specifies from where the shape scales (top-left corner, middle, or bottom-right corner).</span></span> <span data-ttu-id="a31bd-162">プロパティに`lockAspectRatio` **true**が設定されている場合、scale メソッドは、他の次元も調整して、図形の現在の縦横比を維持します。</span><span class="sxs-lookup"><span data-stu-id="a31bd-162">If the `lockAspectRatio` property is **true**, the scale methods maintain the shape's current aspect ratio by also adjusting the other dimension.</span></span>

> [!NOTE]
> <span data-ttu-id="a31bd-163">プロパティに`height` `width`対する直接の変更は、プロパティの値に関係なく`lockAspectRatio` 、そのプロパティにのみ影響します。</span><span class="sxs-lookup"><span data-stu-id="a31bd-163">Direct changes to the `height` and `width` properties only affect that property, regardless of the `lockAspectRatio` property's value.</span></span>

<span data-ttu-id="a31bd-164">次のコードサンプルでは、元のサイズに1.25 倍に拡大または縮小された図形を表示します。</span><span class="sxs-lookup"><span data-stu-id="a31bd-164">The following code sample shows a shape being scaled to 1.25 times its original size and rotated 30 degrees.</span></span>

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

## <a name="text-in-shapes"></a><span data-ttu-id="a31bd-165">図形内のテキスト</span><span class="sxs-lookup"><span data-stu-id="a31bd-165">Text in shapes</span></span>

<span data-ttu-id="a31bd-166">幾何学的な図形にはテキストを含めることができます。</span><span class="sxs-lookup"><span data-stu-id="a31bd-166">Geometric Shapes can contain text.</span></span> <span data-ttu-id="a31bd-167">図形には`textFrame` 、 [TextFrame](/javascript/api/excel/excel.textframe)型のプロパティがあります。</span><span class="sxs-lookup"><span data-stu-id="a31bd-167">Shapes have a `textFrame` property of type [TextFrame](/javascript/api/excel/excel.textframe).</span></span> <span data-ttu-id="a31bd-168">オブジェクト`TextFrame`は、テキスト表示オプション (余白、テキストオーバーフローなど) を管理します。</span><span class="sxs-lookup"><span data-stu-id="a31bd-168">The `TextFrame` object manages the text display options (such as margins and text overflow).</span></span> <span data-ttu-id="a31bd-169">`TextFrame.textRange`は、テキストの内容とフォントの設定を含む[TextRange](/javascript/api/excel/excel.textrange)オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="a31bd-169">`TextFrame.textRange` is a [TextRange](/javascript/api/excel/excel.textrange) object with the text content and font settings.</span></span>

<span data-ttu-id="a31bd-170">次のコードサンプルでは、テキスト "shape text" を使用して "Wave" という名前のジオメトリック図形を作成します。</span><span class="sxs-lookup"><span data-stu-id="a31bd-170">The following code sample creates a geometric shape named "Wave" with the text "Shape Text".</span></span> <span data-ttu-id="a31bd-171">また、図形とテキストの色を調整するだけでなく、テキストの水平方向の配置を中央に設定します。</span><span class="sxs-lookup"><span data-stu-id="a31bd-171">It also adjusts the shape and text colors, as well as sets the text's horizontal alignment to the center.</span></span>

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

<span data-ttu-id="a31bd-172">の`addTextBox` `ShapeCollection`メソッドは、白`GeometricShape`の背景`Rectangle`と黒のテキストを使用して、型を作成します。</span><span class="sxs-lookup"><span data-stu-id="a31bd-172">The `addTextBox` method of `ShapeCollection` creates a `GeometricShape` of type `Rectangle` with a white background and black text.</span></span> <span data-ttu-id="a31bd-173">これは、[挿入] タブの [Excel の**テキストボックス**によって\*\*\*\* 作成され`addTextBox`たものと同じです。文字列型 (string `TextRange`) の引数を指定して、のテキストを設定します。</span><span class="sxs-lookup"><span data-stu-id="a31bd-173">This is the same as what is created by Excel's **Text Box** button on the **Insert** tab. `addTextBox` takes a string argument to set the text of the `TextRange`.</span></span>

<span data-ttu-id="a31bd-174">次のコードサンプルは、"Hello!" というテキストを含むテキストボックスを作成する方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="a31bd-174">The following code sample shows the creation of a text box with the text "Hello!".</span></span>

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

## <a name="shape-groups"></a><span data-ttu-id="a31bd-175">図形グループ</span><span class="sxs-lookup"><span data-stu-id="a31bd-175">Shape groups</span></span>

<span data-ttu-id="a31bd-176">図形は一緒にグループ化できます。</span><span class="sxs-lookup"><span data-stu-id="a31bd-176">Shapes can be grouped together.</span></span> <span data-ttu-id="a31bd-177">これにより、ユーザーは、配置、サイズ変更、およびその他の関連タスクのために1つのエンティティとして扱うことができます。</span><span class="sxs-lookup"><span data-stu-id="a31bd-177">This allows a user to treat them as a single entity for positioning, sizing, and other related tasks.</span></span> <span data-ttu-id="a31bd-178">図形[グループ](/javascript/api/excel/excel.shapegroup)はの`Shape`種類であるため、アドインでグループを1つの図形として扱うことができます。</span><span class="sxs-lookup"><span data-stu-id="a31bd-178">A [ShapeGroup](/javascript/api/excel/excel.shapegroup) is a type of `Shape`, so your add-in treats the group as a single shape.</span></span>

<span data-ttu-id="a31bd-179">次のコードサンプルでは、グループ化された3つの図形を示します。</span><span class="sxs-lookup"><span data-stu-id="a31bd-179">The following code sample shows three shapes being grouped together.</span></span> <span data-ttu-id="a31bd-180">次のコードサンプルでは、図形グループを50ピクセル右に移動していることを示します。</span><span class="sxs-lookup"><span data-stu-id="a31bd-180">The subsequent code sample shows that shape group being moved to the right 50 pixels.</span></span>

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
> <span data-ttu-id="a31bd-181">グループ内の個々の図形は、種類`ShapeGroup.shapes`が[groupshapecollection](/javascript/api/excel/excel.GroupShapeCollection)であるプロパティを介して参照されます。</span><span class="sxs-lookup"><span data-stu-id="a31bd-181">Individual shapes within the group are referenced through the `ShapeGroup.shapes` property, which is of type [GroupShapeCollection](/javascript/api/excel/excel.GroupShapeCollection).</span></span> <span data-ttu-id="a31bd-182">グループ化された後は、ワークシートの shape コレクションからアクセスできなくなります。</span><span class="sxs-lookup"><span data-stu-id="a31bd-182">They are no longer accessible through the worksheet's shape collection after being grouped.</span></span> <span data-ttu-id="a31bd-183">たとえば、ワークシートに3つの図形があり、すべてが一緒にグループ化され`shapes.getCount`ている場合、ワークシートのメソッドはカウントを1とします。</span><span class="sxs-lookup"><span data-stu-id="a31bd-183">As an example, if your worksheet had three shapes and they were all grouped together, the worksheet's `shapes.getCount` method would return a count of 1.</span></span>

## <a name="export-shapes-as-images"></a><span data-ttu-id="a31bd-184">図形を画像としてエクスポートする</span><span class="sxs-lookup"><span data-stu-id="a31bd-184">Export shapes as images</span></span>

<span data-ttu-id="a31bd-185">任意`Shape`のオブジェクトをイメージに変換できます。</span><span class="sxs-lookup"><span data-stu-id="a31bd-185">Any `Shape` object can be converted to an image.</span></span> <span data-ttu-id="a31bd-186">[getAsImage](/javascript/api/excel/excel.shape#getasimage-format-)は、base64 でエンコードされた文字列を返します。</span><span class="sxs-lookup"><span data-stu-id="a31bd-186">[Shape.getAsImage](/javascript/api/excel/excel.shape#getasimage-format-) returns base64-encoded string.</span></span> <span data-ttu-id="a31bd-187">画像の形式は、に`getAsImage`渡される図[形式](/javascript/api/excel/excel.pictureformat)の列挙体として指定されます。</span><span class="sxs-lookup"><span data-stu-id="a31bd-187">The image's format is specified as a [PictureFormat](/javascript/api/excel/excel.pictureformat) enum passed to `getAsImage`.</span></span>

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

## <a name="delete-shapes"></a><span data-ttu-id="a31bd-188">図形を削除する</span><span class="sxs-lookup"><span data-stu-id="a31bd-188">Delete shapes</span></span>

<span data-ttu-id="a31bd-189">図形は、 `Shape`オブジェクトの`delete`メソッドを使用してワークシートから削除されます。</span><span class="sxs-lookup"><span data-stu-id="a31bd-189">Shapes are removed from the worksheet with the `Shape` object's `delete` method.</span></span> <span data-ttu-id="a31bd-190">その他のメタデータは必要ありません。</span><span class="sxs-lookup"><span data-stu-id="a31bd-190">No other metadata is needed.</span></span>

<span data-ttu-id="a31bd-191">次のコードサンプルでは、 **myworksheet**からすべての図形を削除します。</span><span class="sxs-lookup"><span data-stu-id="a31bd-191">The following code sample deletes all the shapes from **MyWorksheet**.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="a31bd-192">関連項目</span><span class="sxs-lookup"><span data-stu-id="a31bd-192">See also</span></span>

- [<span data-ttu-id="a31bd-193">Excel JavaScript API を使用した基本的なプログラミングの概念</span><span class="sxs-lookup"><span data-stu-id="a31bd-193">Fundamental programming concepts with the Excel JavaScript API</span></span>](../reference/overview/excel-add-ins-reference-overview.md)
- [<span data-ttu-id="a31bd-194">Excel JavaScript API を使用してグラフを操作する</span><span class="sxs-lookup"><span data-stu-id="a31bd-194">Work with Charts using the Excel JavaScript API</span></span>](excel-add-ins-charts.md)
