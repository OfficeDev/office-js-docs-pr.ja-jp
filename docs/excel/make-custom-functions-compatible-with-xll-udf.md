---
title: XLL ユーザー定義関数を使用してカスタム関数を拡張する
description: カスタム関数と同等の機能を持つ Excel XLL ユーザー定義関数との互換性を有効にする (プレビュー)
ms.date: 05/06/2019
localization_priority: Normal
ms.openlocfilehash: 93e1b52606fca7ea6fbbb9ae3545e4edd7f78742
ms.sourcegitcommit: ff73cc04e5718765fcbe74181505a974db69c3f5
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/06/2019
ms.locfileid: "33628108"
---
# <a name="extend-custom-functions-with-xll-user-defined-functions-preview"></a><span data-ttu-id="1900a-103">XLL ユーザー定義関数を使用してカスタム関数を拡張する (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="1900a-103">Extend custom functions with XLL user-defined functions (preview)</span></span>

<span data-ttu-id="1900a-104">既存の Excel XLLs がある場合は、Excel アドインで同等のカスタム関数を作成して、online や macOS などの他のプラットフォームにソリューション機能を拡張することができます。</span><span class="sxs-lookup"><span data-stu-id="1900a-104">If you have existing Excel XLLs, you can build equivalent custom functions in an Excel add-in to extend your solution features to other platforms such as online or macOS.</span></span> <span data-ttu-id="1900a-105">ただし、Excel アドインには、xll で利用可能なすべての機能が含まれているわけではありません。</span><span class="sxs-lookup"><span data-stu-id="1900a-105">However, Excel add-ins don't have all of the functionality available in XLLs.</span></span> <span data-ttu-id="1900a-106">ソリューションで使用されている機能によっては、XLL によって excel の excel アドインカスタム関数よりも優れた操作が提供されることがあります。</span><span class="sxs-lookup"><span data-stu-id="1900a-106">Depending on the functionality your solution uses, the XLL may provide a better experience than the Excel add-in custom functions on Excel for Windows.</span></span>

[!include[COM add-in and XLL UDF compatibility note](../includes/xll-compatibility-note.md)]

## <a name="specify-equivalent-xll-in-the-manifest"></a><span data-ttu-id="1900a-107">マニフェストで同等の XLL を指定する</span><span class="sxs-lookup"><span data-stu-id="1900a-107">Specify equivalent XLL in the manifest</span></span>

<span data-ttu-id="1900a-108">既存の XLL との互換性を有効にするには、Excel アドインのマニフェストで同等の XLL を識別します。</span><span class="sxs-lookup"><span data-stu-id="1900a-108">To enable compatibility with an existing XLL, identify the equivalent XLL in the manifest of your Excel add-in.</span></span> <span data-ttu-id="1900a-109">Excel では、Windows での実行時に Excel アドインカスタム関数の代わりに XLL 関数が使用されます。</span><span class="sxs-lookup"><span data-stu-id="1900a-109">Then Excel will use the XLL's functions instead of your Excel add-in custom functions when running on Windows.</span></span>

<span data-ttu-id="1900a-110">カスタム関数に対応する XLL を設定するには、 `FileName` xll のを指定します。</span><span class="sxs-lookup"><span data-stu-id="1900a-110">To set the equivalent XLL for your custom functions, specify the `FileName` of the XLL.</span></span> <span data-ttu-id="1900a-111">ユーザーが XLL から関数を含むブックを開くと、Excel は関数を互換性のある関数に変換します。</span><span class="sxs-lookup"><span data-stu-id="1900a-111">When the user opens a workbook with functions from the XLL, Excel converts the functions to compatible functions.</span></span> <span data-ttu-id="1900a-112">ブックは、Windows の Excel で開いたときに XLL を使用し、オンラインまたは macOS を開いたときに Excel アドインのカスタム関数を使用します。</span><span class="sxs-lookup"><span data-stu-id="1900a-112">The workbook then uses the XLL when opened in Excel on Windows, and it will use custom functions from your Excel add-in when opened online or on macOS.</span></span>

<span data-ttu-id="1900a-113">次の例は、COM アドインと XLL の両方を同等として指定する方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="1900a-113">The following example shows how to specify both a COM add-in and an XLL as equivalent.</span></span> <span data-ttu-id="1900a-114">多くの場合、この例は完全にコンテキストで指定します。</span><span class="sxs-lookup"><span data-stu-id="1900a-114">Often you will specify both so for completeness this example shows both in context.</span></span> <span data-ttu-id="1900a-115">これらは、 `FileName`それぞれに`ProgID`よって識別されます。</span><span class="sxs-lookup"><span data-stu-id="1900a-115">They are identified by their `ProgID` and `FileName` respectively.</span></span> <span data-ttu-id="1900a-116">COM アドインの互換性の詳細については、「[既存の com アドインと互換性のある Excel アドインを作成](../develop/make-office-add-in-compatible-with-existing-com-add-in.md)する」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="1900a-116">For more information on COM add-in compatibility, see [Make your Excel add-in compatible with an existing COM add-in](../develop/make-office-add-in-compatible-with-existing-com-add-in.md).</span></span>

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
> <span data-ttu-id="1900a-117">アドインでカスタム関数が XLL 互換に宣言されている場合、後でマニフェストを変更すると、ファイル形式が変更されるため、ユーザーのブックが破損する可能性があります。</span><span class="sxs-lookup"><span data-stu-id="1900a-117">If an add-in declares its custom functions to be XLL compatible, changing the manifest at a later time could break a user’s workbook because it will change the file format.</span></span>

## <a name="excel-add-in-updates"></a><span data-ttu-id="1900a-118">Excel アドインの更新プログラム</span><span class="sxs-lookup"><span data-stu-id="1900a-118">Excel add-in updates</span></span>

<span data-ttu-id="1900a-119">Excel アドインに対して同等の XLL を指定すると、excel アドインの更新プログラムの処理は中止されます。</span><span class="sxs-lookup"><span data-stu-id="1900a-119">Once you specify an equivalent XLL for your Excel add-in, Excel stops processing updates for your Excel add-in.</span></span> <span data-ttu-id="1900a-120">ユーザーは、Excel アドインの最新の更新プログラムを取得するために、XLL をアンインストールする必要があります。</span><span class="sxs-lookup"><span data-stu-id="1900a-120">The user must uninstall the XLL in order to get the latest updates for the Excel add-in.</span></span>

## <a name="custom-function-behavior-for-xll-compatible-functions"></a><span data-ttu-id="1900a-121">XLL 互換関数のカスタム関数の動作</span><span class="sxs-lookup"><span data-stu-id="1900a-121">Custom function behavior for XLL compatible functions</span></span>

<span data-ttu-id="1900a-122">同じアドインが含まれている XLL 関数を含むスプレッドシートが開かれると、xll 関数は、XLL 互換のカスタム関数に変換されます。</span><span class="sxs-lookup"><span data-stu-id="1900a-122">When a spreadsheet is opened that contains XLL functions for which there is also an equivalent add-in, the XLL's functions are converted to XLL compatible custom functions.</span></span> <span data-ttu-id="1900a-123">次の保存時に、これらのファイルは互換モードでファイルに書き込まれます。これにより、(他のプラットフォームでの場合) XLL と Excel アドインの両方のカスタム機能を使用できるようになります。</span><span class="sxs-lookup"><span data-stu-id="1900a-123">On the next save, they are written to the file in a compatible mode so that they work with both the XLL and Excel add-in custom functions (when on other platforms).</span></span>

<span data-ttu-id="1900a-124">次の表は、XLL ユーザー定義関数、XLL 互換カスタム関数、および Excel アドインカスタム関数の機能を比較しています。</span><span class="sxs-lookup"><span data-stu-id="1900a-124">The following table compares features across XLL user-defined functions, XLL compatible custom functions, and Excel add-in custom functions.</span></span>

|         |<span data-ttu-id="1900a-125">XLL のユーザー定義関数</span><span class="sxs-lookup"><span data-stu-id="1900a-125">XLL user-defined function</span></span> |<span data-ttu-id="1900a-126">XLL 互換のカスタム関数</span><span class="sxs-lookup"><span data-stu-id="1900a-126">XLL compatible custom functions</span></span> |<span data-ttu-id="1900a-127">Excel アドインのカスタム関数</span><span class="sxs-lookup"><span data-stu-id="1900a-127">Excel add-in custom function</span></span> |
|---------|---------|---------|---------|
| <span data-ttu-id="1900a-128">サポートされるプラットフォーム</span><span class="sxs-lookup"><span data-stu-id="1900a-128">Supported platforms</span></span> | <span data-ttu-id="1900a-129">Windows</span><span class="sxs-lookup"><span data-stu-id="1900a-129">Windows</span></span> | <span data-ttu-id="1900a-130">Windows、macOS、Excel online</span><span class="sxs-lookup"><span data-stu-id="1900a-130">Windows, macOS, Excel online</span></span> | <span data-ttu-id="1900a-131">Windows、macOS、Excel online</span><span class="sxs-lookup"><span data-stu-id="1900a-131">Windows, macOS, Excel online</span></span> |
| <span data-ttu-id="1900a-132">サポートされるファイル形式</span><span class="sxs-lookup"><span data-stu-id="1900a-132">Supported file formats</span></span> | <span data-ttu-id="1900a-133">.XLSX、.XLSB、.XLSM、XLS</span><span class="sxs-lookup"><span data-stu-id="1900a-133">XLSX, XLSB, XLSM, XLS</span></span> | <span data-ttu-id="1900a-134">.XLSX、.XLSB、.XLSM</span><span class="sxs-lookup"><span data-stu-id="1900a-134">XLSX, XLSB, XLSM</span></span> | <span data-ttu-id="1900a-135">.XLSX、.XLSB、.XLSM</span><span class="sxs-lookup"><span data-stu-id="1900a-135">XLSX, XLSB, XLSM</span></span> |
| <span data-ttu-id="1900a-136">数式オートコンプリート</span><span class="sxs-lookup"><span data-stu-id="1900a-136">Formula autocomplete</span></span> | <span data-ttu-id="1900a-137">いいえ</span><span class="sxs-lookup"><span data-stu-id="1900a-137">No</span></span> | <span data-ttu-id="1900a-138">はい</span><span class="sxs-lookup"><span data-stu-id="1900a-138">Yes</span></span> | <span data-ttu-id="1900a-139">はい</span><span class="sxs-lookup"><span data-stu-id="1900a-139">Yes</span></span> |
| <span data-ttu-id="1900a-140">ストリーミング</span><span class="sxs-lookup"><span data-stu-id="1900a-140">Streaming</span></span> | <span data-ttu-id="1900a-141">XlfRTD および XLL コールバックを使用して可能。</span><span class="sxs-lookup"><span data-stu-id="1900a-141">Possible via xlfRTD and XLL callback.</span></span> | <span data-ttu-id="1900a-142">いいえ</span><span class="sxs-lookup"><span data-stu-id="1900a-142">No</span></span> | <span data-ttu-id="1900a-143">はい</span><span class="sxs-lookup"><span data-stu-id="1900a-143">Yes</span></span> |
| <span data-ttu-id="1900a-144">関数のローカライズ</span><span class="sxs-lookup"><span data-stu-id="1900a-144">Localization of functions</span></span> | <span data-ttu-id="1900a-145">いいえ</span><span class="sxs-lookup"><span data-stu-id="1900a-145">No</span></span> | <span data-ttu-id="1900a-146">いいえ。</span><span class="sxs-lookup"><span data-stu-id="1900a-146">No.</span></span> <span data-ttu-id="1900a-147">名前と ID は、既存の XLL 関数と一致している必要があります。</span><span class="sxs-lookup"><span data-stu-id="1900a-147">The Name and ID must match the existing XLL's functions.</span></span> | <span data-ttu-id="1900a-148">はい</span><span class="sxs-lookup"><span data-stu-id="1900a-148">Yes</span></span> |
| <span data-ttu-id="1900a-149">揮発性関数</span><span class="sxs-lookup"><span data-stu-id="1900a-149">Volatile functions</span></span> | <span data-ttu-id="1900a-150">はい</span><span class="sxs-lookup"><span data-stu-id="1900a-150">Yes</span></span> | <span data-ttu-id="1900a-151">はい</span><span class="sxs-lookup"><span data-stu-id="1900a-151">Yes</span></span> | <span data-ttu-id="1900a-152">はい</span><span class="sxs-lookup"><span data-stu-id="1900a-152">Yes</span></span> |
| <span data-ttu-id="1900a-153">マルチスレッドの再計算のサポート</span><span class="sxs-lookup"><span data-stu-id="1900a-153">Multi-threaded recalculation support</span></span> | <span data-ttu-id="1900a-154">はい</span><span class="sxs-lookup"><span data-stu-id="1900a-154">Yes</span></span> | <span data-ttu-id="1900a-155">はい</span><span class="sxs-lookup"><span data-stu-id="1900a-155">Yes</span></span> | <span data-ttu-id="1900a-156">はい</span><span class="sxs-lookup"><span data-stu-id="1900a-156">Yes</span></span> |
| <span data-ttu-id="1900a-157">計算動作</span><span class="sxs-lookup"><span data-stu-id="1900a-157">Calculation behavior</span></span> | <span data-ttu-id="1900a-158">UI がありません。</span><span class="sxs-lookup"><span data-stu-id="1900a-158">No UI.</span></span> <span data-ttu-id="1900a-159">計算中に Excel が応答しなくなることがあります。</span><span class="sxs-lookup"><span data-stu-id="1900a-159">Excel can be unresponsive during calculation.</span></span> | <span data-ttu-id="1900a-160">ユーザーには #BUSY が表示されます。</span><span class="sxs-lookup"><span data-stu-id="1900a-160">Users will see #BUSY!</span></span> <span data-ttu-id="1900a-161">を返します。</span><span class="sxs-lookup"><span data-stu-id="1900a-161">until a result is returned.</span></span> | <span data-ttu-id="1900a-162">ユーザーには #BUSY が表示されます。</span><span class="sxs-lookup"><span data-stu-id="1900a-162">Users will see #BUSY!</span></span> <span data-ttu-id="1900a-163">を返します。</span><span class="sxs-lookup"><span data-stu-id="1900a-163">until a result is returned.</span></span> |
| <span data-ttu-id="1900a-164">要件セット</span><span class="sxs-lookup"><span data-stu-id="1900a-164">Requirement sets</span></span> | <span data-ttu-id="1900a-165">N/A</span><span class="sxs-lookup"><span data-stu-id="1900a-165">N/A</span></span> | <span data-ttu-id="1900a-166">CustomFunctions 1.1 以降</span><span class="sxs-lookup"><span data-stu-id="1900a-166">CustomFunctions 1.1 and later</span></span> | <span data-ttu-id="1900a-167">CustomFunctions 1.1 以降</span><span class="sxs-lookup"><span data-stu-id="1900a-167">CustomFunctions 1.1 and later</span></span> |

## <a name="see-also"></a><span data-ttu-id="1900a-168">関連項目</span><span class="sxs-lookup"><span data-stu-id="1900a-168">See also</span></span>

- [<span data-ttu-id="1900a-169">既存の COM アドインと互換性のある Excel アドインを作成する</span><span class="sxs-lookup"><span data-stu-id="1900a-169">Make your Excel add-in compatible with an existing COM add-in</span></span>](../develop/make-office-add-in-compatible-with-existing-com-add-in.md)
- [<span data-ttu-id="1900a-170">カスタム関数のベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="1900a-170">Custom functions best practices</span></span>](custom-functions-best-practices.md)
- [<span data-ttu-id="1900a-171">チュートリアル: Excel でカスタム関数を作成します。</span><span class="sxs-lookup"><span data-stu-id="1900a-171">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
