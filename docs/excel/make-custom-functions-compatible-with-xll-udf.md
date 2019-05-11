---
title: XLL ユーザー定義関数を使用してカスタム関数を拡張する
description: カスタム関数と同等の機能を持つ Excel XLL ユーザー定義関数との互換性を有効にする (プレビュー)
ms.date: 05/08/2019
localization_priority: Normal
ms.openlocfilehash: 3e1782c5df227d3e173f4291ba88f2057200b1c5
ms.sourcegitcommit: a99be9c4771c45f3e07e781646e0e649aa47213f
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/11/2019
ms.locfileid: "33951887"
---
# <a name="extend-custom-functions-with-xll-user-defined-functions-preview"></a><span data-ttu-id="1d69c-103">XLL ユーザー定義関数を使用してカスタム関数を拡張する (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="1d69c-103">Extend custom functions with XLL user-defined functions (preview)</span></span>

<span data-ttu-id="1d69c-104">既存の Excel XLLs がある場合は、Excel アドインで同等のカスタム関数を作成して、online や macOS などの他のプラットフォームにソリューション機能を拡張することができます。</span><span class="sxs-lookup"><span data-stu-id="1d69c-104">If you have existing Excel XLLs, you can build equivalent custom functions in an Excel add-in to extend your solution features to other platforms such as online or macOS.</span></span> <span data-ttu-id="1d69c-105">ただし、Excel アドインには、xll で利用可能なすべての機能が含まれているわけではありません。</span><span class="sxs-lookup"><span data-stu-id="1d69c-105">However, Excel add-ins don't have all of the functionality available in XLLs.</span></span> <span data-ttu-id="1d69c-106">ソリューションで使用されている機能によっては、XLL の方が excel の excel アドインカスタム関数よりも優れた操作を提供することがあります。</span><span class="sxs-lookup"><span data-stu-id="1d69c-106">Depending on the functionality your solution uses, the XLL may provide a better experience than the Excel add-in custom functions in Excel on Windows.</span></span>

[!include[COM add-in and XLL UDF compatibility note](../includes/xll-compatibility-note.md)]

## <a name="specify-equivalent-xll-in-the-manifest"></a><span data-ttu-id="1d69c-107">マニフェストで同等の XLL を指定する</span><span class="sxs-lookup"><span data-stu-id="1d69c-107">Specify equivalent XLL in the manifest</span></span>

<span data-ttu-id="1d69c-108">既存の XLL との互換性を有効にするには、Excel アドインのマニフェストで同等の XLL を識別します。</span><span class="sxs-lookup"><span data-stu-id="1d69c-108">To enable compatibility with an existing XLL, identify the equivalent XLL in the manifest of your Excel add-in.</span></span> <span data-ttu-id="1d69c-109">Excel では、Windows での実行時に Excel アドインカスタム関数の代わりに XLL 関数が使用されます。</span><span class="sxs-lookup"><span data-stu-id="1d69c-109">Then Excel will use the XLL's functions instead of your Excel add-in custom functions when running on Windows.</span></span>

<span data-ttu-id="1d69c-110">カスタム関数に対応する XLL を設定するには、 `FileName` xll のを指定します。</span><span class="sxs-lookup"><span data-stu-id="1d69c-110">To set the equivalent XLL for your custom functions, specify the `FileName` of the XLL.</span></span> <span data-ttu-id="1d69c-111">ユーザーが XLL から関数を含むブックを開くと、Excel は関数を互換性のある関数に変換します。</span><span class="sxs-lookup"><span data-stu-id="1d69c-111">When the user opens a workbook with functions from the XLL, Excel converts the functions to compatible functions.</span></span> <span data-ttu-id="1d69c-112">ブックは、Windows の Excel で開いたときに XLL を使用し、オンラインまたは macOS を開いたときに Excel アドインのカスタム関数を使用します。</span><span class="sxs-lookup"><span data-stu-id="1d69c-112">The workbook then uses the XLL when opened in Excel on Windows, and it will use custom functions from your Excel add-in when opened online or on macOS.</span></span>

<span data-ttu-id="1d69c-113">次の例は、COM アドインと XLL の両方を同等として指定する方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="1d69c-113">The following example shows how to specify both a COM add-in and an XLL as equivalent.</span></span> <span data-ttu-id="1d69c-114">多くの場合、この例は完全にコンテキストで指定します。</span><span class="sxs-lookup"><span data-stu-id="1d69c-114">Often you will specify both so for completeness this example shows both in context.</span></span> <span data-ttu-id="1d69c-115">これらは、 `FileName`それぞれに`ProgID`よって識別されます。</span><span class="sxs-lookup"><span data-stu-id="1d69c-115">They are identified by their `ProgID` and `FileName` respectively.</span></span> <span data-ttu-id="1d69c-116">COM アドインの互換性の詳細については、「[既存の com アドインと互換性のある Excel アドインを作成](../develop/make-office-add-in-compatible-with-existing-com-add-in.md)する」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="1d69c-116">For more information on COM add-in compatibility, see [Make your Excel add-in compatible with an existing COM add-in](../develop/make-office-add-in-compatible-with-existing-com-add-in.md).</span></span>

```xml
<VersionOverrides>
...
<EquivalentAddins>
  <EquivalentAddin>
    <ProgID>ContosoCOMAddin</ProgID>
    <Type>COM</Type>
  </EquivalentAddin>

  <EquivalentAddin>
    <FileName>contosofunctions.xll</FileName>
    <Type>XLL</Type>
  </EquivalentAddin>
<EquivalentAddins>
...
</VersionOverrides>
```

> [!NOTE]
> <span data-ttu-id="1d69c-117">アドインでカスタム関数が XLL 互換に宣言されている場合、後でマニフェストを変更すると、ファイル形式が変更されるため、ユーザーのブックが破損する可能性があります。</span><span class="sxs-lookup"><span data-stu-id="1d69c-117">If an add-in declares its custom functions to be XLL compatible, changing the manifest at a later time could break a user’s workbook because it will change the file format.</span></span>

## <a name="excel-add-in-updates"></a><span data-ttu-id="1d69c-118">Excel アドインの更新プログラム</span><span class="sxs-lookup"><span data-stu-id="1d69c-118">Excel add-in updates</span></span>

<span data-ttu-id="1d69c-119">Excel アドインに対して同等の XLL を指定すると、excel アドインの更新プログラムの処理は中止されます。</span><span class="sxs-lookup"><span data-stu-id="1d69c-119">Once you specify an equivalent XLL for your Excel add-in, Excel stops processing updates for your Excel add-in.</span></span> <span data-ttu-id="1d69c-120">ユーザーは、Excel アドインの最新の更新プログラムを取得するために、XLL をアンインストールする必要があります。</span><span class="sxs-lookup"><span data-stu-id="1d69c-120">The user must uninstall the XLL in order to get the latest updates for the Excel add-in.</span></span>

## <a name="custom-function-behavior-for-xll-compatible-functions"></a><span data-ttu-id="1d69c-121">XLL 互換関数のカスタム関数の動作</span><span class="sxs-lookup"><span data-stu-id="1d69c-121">Custom function behavior for XLL compatible functions</span></span>

<span data-ttu-id="1d69c-122">同じアドインが含まれている XLL 関数を含むスプレッドシートが開かれると、xll 関数は、XLL 互換のカスタム関数に変換されます。</span><span class="sxs-lookup"><span data-stu-id="1d69c-122">When a spreadsheet is opened that contains XLL functions for which there is also an equivalent add-in, the XLL's functions are converted to XLL compatible custom functions.</span></span> <span data-ttu-id="1d69c-123">次の保存時に、これらのファイルは互換モードでファイルに書き込まれます。これにより、(他のプラットフォームでの場合) XLL と Excel アドインの両方のカスタム機能を使用できるようになります。</span><span class="sxs-lookup"><span data-stu-id="1d69c-123">On the next save, they are written to the file in a compatible mode so that they work with both the XLL and Excel add-in custom functions (when on other platforms).</span></span>

<span data-ttu-id="1d69c-124">次の表は、XLL ユーザー定義関数、XLL 互換カスタム関数、および Excel アドインカスタム関数の機能を比較しています。</span><span class="sxs-lookup"><span data-stu-id="1d69c-124">The following table compares features across XLL user-defined functions, XLL compatible custom functions, and Excel add-in custom functions.</span></span>

|         |<span data-ttu-id="1d69c-125">XLL のユーザー定義関数</span><span class="sxs-lookup"><span data-stu-id="1d69c-125">XLL user-defined function</span></span> |<span data-ttu-id="1d69c-126">XLL 互換のカスタム関数</span><span class="sxs-lookup"><span data-stu-id="1d69c-126">XLL compatible custom functions</span></span> |<span data-ttu-id="1d69c-127">Excel アドインのカスタム関数</span><span class="sxs-lookup"><span data-stu-id="1d69c-127">Excel add-in custom function</span></span> |
|---------|---------|---------|---------|
| <span data-ttu-id="1d69c-128">サポートされるプラットフォーム</span><span class="sxs-lookup"><span data-stu-id="1d69c-128">Supported platforms</span></span> | <span data-ttu-id="1d69c-129">Windows</span><span class="sxs-lookup"><span data-stu-id="1d69c-129">Windows</span></span> | <span data-ttu-id="1d69c-130">Windows、macOS、Excel online</span><span class="sxs-lookup"><span data-stu-id="1d69c-130">Windows, macOS, Excel online</span></span> | <span data-ttu-id="1d69c-131">Windows、macOS、Excel online</span><span class="sxs-lookup"><span data-stu-id="1d69c-131">Windows, macOS, Excel online</span></span> |
| <span data-ttu-id="1d69c-132">サポートされるファイル形式</span><span class="sxs-lookup"><span data-stu-id="1d69c-132">Supported file formats</span></span> | <span data-ttu-id="1d69c-133">.XLSX、.XLSB、.XLSM、XLS</span><span class="sxs-lookup"><span data-stu-id="1d69c-133">XLSX, XLSB, XLSM, XLS</span></span> | <span data-ttu-id="1d69c-134">.XLSX、.XLSB、.XLSM</span><span class="sxs-lookup"><span data-stu-id="1d69c-134">XLSX, XLSB, XLSM</span></span> | <span data-ttu-id="1d69c-135">.XLSX、.XLSB、.XLSM</span><span class="sxs-lookup"><span data-stu-id="1d69c-135">XLSX, XLSB, XLSM</span></span> |
| <span data-ttu-id="1d69c-136">数式オートコンプリート</span><span class="sxs-lookup"><span data-stu-id="1d69c-136">Formula autocomplete</span></span> | <span data-ttu-id="1d69c-137">いいえ</span><span class="sxs-lookup"><span data-stu-id="1d69c-137">No</span></span> | <span data-ttu-id="1d69c-138">はい</span><span class="sxs-lookup"><span data-stu-id="1d69c-138">Yes</span></span> | <span data-ttu-id="1d69c-139">はい</span><span class="sxs-lookup"><span data-stu-id="1d69c-139">Yes</span></span> |
| <span data-ttu-id="1d69c-140">ストリーミング</span><span class="sxs-lookup"><span data-stu-id="1d69c-140">Streaming</span></span> | <span data-ttu-id="1d69c-141">XlfRTD および XLL コールバックを使用して可能。</span><span class="sxs-lookup"><span data-stu-id="1d69c-141">Possible via xlfRTD and XLL callback.</span></span> | <span data-ttu-id="1d69c-142">いいえ</span><span class="sxs-lookup"><span data-stu-id="1d69c-142">No</span></span> | <span data-ttu-id="1d69c-143">はい</span><span class="sxs-lookup"><span data-stu-id="1d69c-143">Yes</span></span> |
| <span data-ttu-id="1d69c-144">関数のローカライズ</span><span class="sxs-lookup"><span data-stu-id="1d69c-144">Localization of functions</span></span> | <span data-ttu-id="1d69c-145">不要</span><span class="sxs-lookup"><span data-stu-id="1d69c-145">No</span></span> | <span data-ttu-id="1d69c-146">いいえ。</span><span class="sxs-lookup"><span data-stu-id="1d69c-146">No.</span></span> <span data-ttu-id="1d69c-147">名前と ID は、既存の XLL 関数と一致している必要があります。</span><span class="sxs-lookup"><span data-stu-id="1d69c-147">The Name and ID must match the existing XLL's functions.</span></span> | <span data-ttu-id="1d69c-148">はい</span><span class="sxs-lookup"><span data-stu-id="1d69c-148">Yes</span></span> |
| <span data-ttu-id="1d69c-149">揮発性関数</span><span class="sxs-lookup"><span data-stu-id="1d69c-149">Volatile functions</span></span> | <span data-ttu-id="1d69c-150">はい</span><span class="sxs-lookup"><span data-stu-id="1d69c-150">Yes</span></span> | <span data-ttu-id="1d69c-151">はい</span><span class="sxs-lookup"><span data-stu-id="1d69c-151">Yes</span></span> | <span data-ttu-id="1d69c-152">はい</span><span class="sxs-lookup"><span data-stu-id="1d69c-152">Yes</span></span> |
| <span data-ttu-id="1d69c-153">マルチスレッドの再計算のサポート</span><span class="sxs-lookup"><span data-stu-id="1d69c-153">Multi-threaded recalculation support</span></span> | <span data-ttu-id="1d69c-154">はい</span><span class="sxs-lookup"><span data-stu-id="1d69c-154">Yes</span></span> | <span data-ttu-id="1d69c-155">はい</span><span class="sxs-lookup"><span data-stu-id="1d69c-155">Yes</span></span> | <span data-ttu-id="1d69c-156">はい</span><span class="sxs-lookup"><span data-stu-id="1d69c-156">Yes</span></span> |
| <span data-ttu-id="1d69c-157">計算動作</span><span class="sxs-lookup"><span data-stu-id="1d69c-157">Calculation behavior</span></span> | <span data-ttu-id="1d69c-158">UI がありません。</span><span class="sxs-lookup"><span data-stu-id="1d69c-158">No UI.</span></span> <span data-ttu-id="1d69c-159">計算中に Excel が応答しなくなることがあります。</span><span class="sxs-lookup"><span data-stu-id="1d69c-159">Excel can be unresponsive during calculation.</span></span> | <span data-ttu-id="1d69c-160">ユーザーには #BUSY が表示されます。</span><span class="sxs-lookup"><span data-stu-id="1d69c-160">Users will see #BUSY!</span></span> <span data-ttu-id="1d69c-161">を返します。</span><span class="sxs-lookup"><span data-stu-id="1d69c-161">until a result is returned.</span></span> | <span data-ttu-id="1d69c-162">ユーザーには #BUSY が表示されます。</span><span class="sxs-lookup"><span data-stu-id="1d69c-162">Users will see #BUSY!</span></span> <span data-ttu-id="1d69c-163">を返します。</span><span class="sxs-lookup"><span data-stu-id="1d69c-163">until a result is returned.</span></span> |
| <span data-ttu-id="1d69c-164">要件セット</span><span class="sxs-lookup"><span data-stu-id="1d69c-164">Requirement sets</span></span> | <span data-ttu-id="1d69c-165">N/A</span><span class="sxs-lookup"><span data-stu-id="1d69c-165">N/A</span></span> | <span data-ttu-id="1d69c-166">CustomFunctions 1.1 以降</span><span class="sxs-lookup"><span data-stu-id="1d69c-166">CustomFunctions 1.1 and later</span></span> | <span data-ttu-id="1d69c-167">CustomFunctions 1.1 以降</span><span class="sxs-lookup"><span data-stu-id="1d69c-167">CustomFunctions 1.1 and later</span></span> |

## <a name="see-also"></a><span data-ttu-id="1d69c-168">関連項目</span><span class="sxs-lookup"><span data-stu-id="1d69c-168">See also</span></span>

- [<span data-ttu-id="1d69c-169">既存の COM アドインと互換性のある Excel アドインを作成する</span><span class="sxs-lookup"><span data-stu-id="1d69c-169">Make your Excel add-in compatible with an existing COM add-in</span></span>](../develop/make-office-add-in-compatible-with-existing-com-add-in.md)
- [<span data-ttu-id="1d69c-170">カスタム関数のベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="1d69c-170">Custom functions best practices</span></span>](custom-functions-best-practices.md)
- [<span data-ttu-id="1d69c-171">Excel カスタム関数のチュートリアル</span><span class="sxs-lookup"><span data-stu-id="1d69c-171">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
