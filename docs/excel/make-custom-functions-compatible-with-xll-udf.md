---
title: カスタム関数を XLL ユーザー定義関数と互換性があるようにする
description: カスタム関数と同等の機能を持つ Excel XLL ユーザー定義関数との互換性を有効にする
ms.date: 04/22/2019
localization_priority: Normal
ms.openlocfilehash: 09914e040c1721dd8b9e91952e5814e7a6b914e5
ms.sourcegitcommit: 7462409209264dc7f8f89f3808a7a6249fcd739e
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/26/2019
ms.locfileid: "33356897"
---
# <a name="make-your-custom-functions-compatible-with-xll-user-defined-functions"></a><span data-ttu-id="f96d8-103">カスタム関数を XLL ユーザー定義関数と互換性があるようにする</span><span class="sxs-lookup"><span data-stu-id="f96d8-103">Make your custom functions compatible with XLL user-defined functions</span></span>

<span data-ttu-id="f96d8-104">既存の Excel xlls がある場合は、Office アドインで同等のカスタム関数を構築して、online や macOS などの他のプラットフォームにソリューション機能を拡張することができます。</span><span class="sxs-lookup"><span data-stu-id="f96d8-104">If you have existing Excel XLLs, you can build equivalent custom functions in an Office Add-in to extend your solution features to other platforms such as online or macOS.</span></span> <span data-ttu-id="f96d8-105">ただし、Office アドインには、すべての機能が xlls で利用できるわけではありません。</span><span class="sxs-lookup"><span data-stu-id="f96d8-105">However, Office Add-ins don't have all of the functionality available in XLLs.</span></span> <span data-ttu-id="f96d8-106">ソリューションで使用されている機能によっては、XLL によって、Excel for Windows の Office アドインカスタム関数よりも優れた操作が提供されることがあります。</span><span class="sxs-lookup"><span data-stu-id="f96d8-106">Depending on the functionality your solution uses, the XLL may provide a better experience than the Office Add-in custom functions on Excel for Windows.</span></span>

<span data-ttu-id="f96d8-107">同等の xll がユーザーのコンピューターに既にインストールされている場合は、office アドインのカスタム関数ではなく、xll が実行されるように、office アドインを構成することができます。</span><span class="sxs-lookup"><span data-stu-id="f96d8-107">You can configure your Office Add-in so that when an equivalent XLL is already installed on the user's computer, Excel runs the XLL instead of your Office Add-in custom functions.</span></span> <span data-ttu-id="f96d8-108">Excel では、Windows にインストールされているを基にして xll と Office アドインカスタム関数の間で切り替えがシームレスに行われるため、xll は同等と呼ばれます。</span><span class="sxs-lookup"><span data-stu-id="f96d8-108">The XLL is called equivalent because Excel will seamlessly transition between the XLL and the Office Add-in custom functions depending on which is installed on Windows.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="specify-equivalent-xll-in-the-manifest"></a><span data-ttu-id="f96d8-109">マニフェストで同等の XLL を指定する</span><span class="sxs-lookup"><span data-stu-id="f96d8-109">Specify equivalent XLL in the manifest</span></span>

<span data-ttu-id="f96d8-110">既存の xll との互換性を有効にするには、Office アドインのマニフェストで同等の xll を識別します。</span><span class="sxs-lookup"><span data-stu-id="f96d8-110">To enable compatibility with an existing XLL, identify the equivalent XLL in the manifest of your Office Add-in.</span></span> <span data-ttu-id="f96d8-111">その後、Excel での実行時に、Office アドインカスタム関数の代わりに XLL 関数が使用されます。</span><span class="sxs-lookup"><span data-stu-id="f96d8-111">Then Excel will use the XLL's functions instead of your Office Add-in custom functions when running on Windows.</span></span>

<span data-ttu-id="f96d8-112">カスタム関数に対応する xll を設定するには、 `FileName` xll のを指定します。</span><span class="sxs-lookup"><span data-stu-id="f96d8-112">To set the equivalent XLL for your custom functions, specify the `FileName` of the XLL.</span></span> <span data-ttu-id="f96d8-113">ユーザーが XLL から関数を含むブックを開くと、Excel は関数を互換性のある関数に変換します。</span><span class="sxs-lookup"><span data-stu-id="f96d8-113">When the user opens a workbook with functions from the XLL, Excel converts the functions to compatible functions.</span></span> <span data-ttu-id="f96d8-114">ブックは、Windows 上の Excel で開いたときに XLL を使用し、オンラインまたは macOS で開いたときに Office アドインのカスタム関数を使用するようになります。</span><span class="sxs-lookup"><span data-stu-id="f96d8-114">The workbook then uses the XLL when opened in Excel on Windows, and it will use custom functions from your Office Add-in when opened online or on macOS.</span></span>

<span data-ttu-id="f96d8-115">次の例は、COM アドインと XLL の両方を同等として指定する方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="f96d8-115">The following example shows how to specify both a COM add-in and an XLL as equivalent.</span></span> <span data-ttu-id="f96d8-116">多くの場合、この例は完全にコンテキストで指定します。</span><span class="sxs-lookup"><span data-stu-id="f96d8-116">Often you will specify both so for completeness this example shows both in context.</span></span> <span data-ttu-id="f96d8-117">これらは、 `FileName`それぞれに`ProgID`よって識別されます。</span><span class="sxs-lookup"><span data-stu-id="f96d8-117">They are identified by their `ProgID` and `FileName` respectively.</span></span> <span data-ttu-id="f96d8-118">com アドインの互換性の詳細については、「[既存の com アドインと互換性のある Office アドインを作成](../develop/make-office-add-in-compatible-with-existing-com-add-in.md)する」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="f96d8-118">For more information on COM add-in compatibility, see [Make your Office Add-in compatible with an existing COM add-in](../develop/make-office-add-in-compatible-with-existing-com-add-in.md).</span></span>

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
> <span data-ttu-id="f96d8-119">アドインでカスタム関数が XLL 互換に宣言されている場合、後でマニフェストを変更すると、ファイル形式が変更されるため、ユーザーのブックが破損する可能性があります。</span><span class="sxs-lookup"><span data-stu-id="f96d8-119">If an add-in declares its custom functions to be XLL compatible, changing the manifest at a later time could break a user’s workbook because it will change the file format.</span></span>

## <a name="office-add-in-updates"></a><span data-ttu-id="f96d8-120">Office アドインの更新プログラム</span><span class="sxs-lookup"><span data-stu-id="f96d8-120">Office Add-in updates</span></span>

<span data-ttu-id="f96d8-121">office アドインに対して同等の XLL を指定すると、Excel は office アドインの更新プログラムの処理を停止します。</span><span class="sxs-lookup"><span data-stu-id="f96d8-121">Once you specify an equivalent XLL for your Office Add-in, Excel stops processing updates for your Office Add-in.</span></span> <span data-ttu-id="f96d8-122">ユーザーは、Office アドインの最新の更新プログラムを入手するために XLL をアンインストールする必要があります。</span><span class="sxs-lookup"><span data-stu-id="f96d8-122">The user must uninstall the XLL in order to get the latest updates for the Office Add-in.</span></span>

## <a name="custom-function-behavior-for-xll-compatible-functions"></a><span data-ttu-id="f96d8-123">XLL 互換関数のカスタム関数の動作</span><span class="sxs-lookup"><span data-stu-id="f96d8-123">Custom function behavior for XLL compatible functions</span></span>

<span data-ttu-id="f96d8-124">同じアドインが含まれている xll 関数を含むスプレッドシートが開かれると、xll 関数は、xll 互換のカスタム関数に変換されます。</span><span class="sxs-lookup"><span data-stu-id="f96d8-124">When a spreadsheet is opened that contains XLL functions for which there is also an equivalent add-in, the XLL's functions are converted to XLL compatible custom functions.</span></span> <span data-ttu-id="f96d8-125">次の保存時に、これらのファイルは互換モードでファイルに書き込まれ、XLL と Office アドインのカスタム関数 (他のプラットフォーム上の場合) の両方で動作するようになります。</span><span class="sxs-lookup"><span data-stu-id="f96d8-125">On the next save, they are written to the file in a compatible mode so that they work with both the XLL and Office Add-in custom functions (when on other platforms).</span></span>

<span data-ttu-id="f96d8-126">次の表は、xll ユーザー定義関数、xll 互換カスタム関数、Office アドインカスタム関数の機能を比較しています。</span><span class="sxs-lookup"><span data-stu-id="f96d8-126">The following table compares features across XLL user-defined functions, XLL compatible custom functions, and Office Add-in custom functions.</span></span>

|         |<span data-ttu-id="f96d8-127">XLL のユーザー定義関数</span><span class="sxs-lookup"><span data-stu-id="f96d8-127">XLL user-defined function</span></span> |<span data-ttu-id="f96d8-128">XLL 互換のカスタム関数</span><span class="sxs-lookup"><span data-stu-id="f96d8-128">XLL compatible custom functions</span></span> |<span data-ttu-id="f96d8-129">Office アドインカスタム関数</span><span class="sxs-lookup"><span data-stu-id="f96d8-129">Office Add-in custom function</span></span> |
|---------|---------|---------|---------|
| <span data-ttu-id="f96d8-130">サポートされるプラットフォーム</span><span class="sxs-lookup"><span data-stu-id="f96d8-130">Supported platforms</span></span> | <span data-ttu-id="f96d8-131">Windows</span><span class="sxs-lookup"><span data-stu-id="f96d8-131">Windows</span></span> | <span data-ttu-id="f96d8-132">Windows、macOS、Excel online</span><span class="sxs-lookup"><span data-stu-id="f96d8-132">Windows, macOS, Excel online</span></span> | <span data-ttu-id="f96d8-133">Windows、macOS、Excel online</span><span class="sxs-lookup"><span data-stu-id="f96d8-133">Windows, macOS, Excel online</span></span> |
| <span data-ttu-id="f96d8-134">サポートされるファイル形式</span><span class="sxs-lookup"><span data-stu-id="f96d8-134">Supported file formats</span></span> | <span data-ttu-id="f96d8-135">.XLSX、.XLSB、.XLSM、XLS</span><span class="sxs-lookup"><span data-stu-id="f96d8-135">XLSX, XLSB, XLSM, XLS</span></span> | <span data-ttu-id="f96d8-136">.XLSX、.XLSB、.XLSM</span><span class="sxs-lookup"><span data-stu-id="f96d8-136">XLSX, XLSB, XLSM</span></span> | <span data-ttu-id="f96d8-137">.XLSX、.XLSB、.XLSM</span><span class="sxs-lookup"><span data-stu-id="f96d8-137">XLSX, XLSB, XLSM</span></span> |
| <span data-ttu-id="f96d8-138">数式オートコンプリート</span><span class="sxs-lookup"><span data-stu-id="f96d8-138">Formula autocomplete</span></span> | <span data-ttu-id="f96d8-139">いいえ</span><span class="sxs-lookup"><span data-stu-id="f96d8-139">No</span></span> | <span data-ttu-id="f96d8-140">はい</span><span class="sxs-lookup"><span data-stu-id="f96d8-140">Yes</span></span> | <span data-ttu-id="f96d8-141">はい</span><span class="sxs-lookup"><span data-stu-id="f96d8-141">Yes</span></span> |
| <span data-ttu-id="f96d8-142">ストリーミング</span><span class="sxs-lookup"><span data-stu-id="f96d8-142">Streaming</span></span> | <span data-ttu-id="f96d8-143">xlfrtd および XLL コールバックを使用して可能。</span><span class="sxs-lookup"><span data-stu-id="f96d8-143">Possible via xlfRTD and XLL callback.</span></span> | <span data-ttu-id="f96d8-144">はい</span><span class="sxs-lookup"><span data-stu-id="f96d8-144">Yes</span></span> | <span data-ttu-id="f96d8-145">はい</span><span class="sxs-lookup"><span data-stu-id="f96d8-145">Yes</span></span> |
| <span data-ttu-id="f96d8-146">関数のローカライズ</span><span class="sxs-lookup"><span data-stu-id="f96d8-146">Localization of functions</span></span> | <span data-ttu-id="f96d8-147">いいえ</span><span class="sxs-lookup"><span data-stu-id="f96d8-147">No</span></span> | <span data-ttu-id="f96d8-148">いいえ。</span><span class="sxs-lookup"><span data-stu-id="f96d8-148">No.</span></span> <span data-ttu-id="f96d8-149">名前と ID は、既存の XLL 関数と一致している必要があります。</span><span class="sxs-lookup"><span data-stu-id="f96d8-149">The Name and ID must match the existing XLL's functions.</span></span> | <span data-ttu-id="f96d8-150">はい</span><span class="sxs-lookup"><span data-stu-id="f96d8-150">Yes</span></span> |
| <span data-ttu-id="f96d8-151">揮発性関数</span><span class="sxs-lookup"><span data-stu-id="f96d8-151">Volatile functions</span></span> | <span data-ttu-id="f96d8-152">はい</span><span class="sxs-lookup"><span data-stu-id="f96d8-152">Yes</span></span> | <span data-ttu-id="f96d8-153">はい</span><span class="sxs-lookup"><span data-stu-id="f96d8-153">Yes</span></span> | <span data-ttu-id="f96d8-154">はい</span><span class="sxs-lookup"><span data-stu-id="f96d8-154">Yes</span></span> |
| <span data-ttu-id="f96d8-155">マルチスレッドの再計算のサポート</span><span class="sxs-lookup"><span data-stu-id="f96d8-155">Multi-threaded recalculation support</span></span> | <span data-ttu-id="f96d8-156">はい</span><span class="sxs-lookup"><span data-stu-id="f96d8-156">Yes</span></span> | <span data-ttu-id="f96d8-157">はい</span><span class="sxs-lookup"><span data-stu-id="f96d8-157">Yes</span></span> | <span data-ttu-id="f96d8-158">はい</span><span class="sxs-lookup"><span data-stu-id="f96d8-158">Yes</span></span> |
| <span data-ttu-id="f96d8-159">計算動作</span><span class="sxs-lookup"><span data-stu-id="f96d8-159">Calculation behavior</span></span> | <span data-ttu-id="f96d8-160">UI がありません。</span><span class="sxs-lookup"><span data-stu-id="f96d8-160">No UI.</span></span> <span data-ttu-id="f96d8-161">計算中に Excel が応答しなくなることがあります。</span><span class="sxs-lookup"><span data-stu-id="f96d8-161">Excel can be unresponsive during calculation.</span></span> | <span data-ttu-id="f96d8-162">ユーザーには #BUSY が表示されます。</span><span class="sxs-lookup"><span data-stu-id="f96d8-162">Users will see #BUSY!</span></span> <span data-ttu-id="f96d8-163">を返します。</span><span class="sxs-lookup"><span data-stu-id="f96d8-163">until a result is returned.</span></span> | <span data-ttu-id="f96d8-164">ユーザーには #BUSY が表示されます。</span><span class="sxs-lookup"><span data-stu-id="f96d8-164">Users will see #BUSY!</span></span> <span data-ttu-id="f96d8-165">を返します。</span><span class="sxs-lookup"><span data-stu-id="f96d8-165">until a result is returned.</span></span> |
| <span data-ttu-id="f96d8-166">要件セット</span><span class="sxs-lookup"><span data-stu-id="f96d8-166">Requirement sets</span></span> | <span data-ttu-id="f96d8-167">該当なし</span><span class="sxs-lookup"><span data-stu-id="f96d8-167">N/A</span></span> | <span data-ttu-id="f96d8-168">customfunctions 1.1 のみ</span><span class="sxs-lookup"><span data-stu-id="f96d8-168">CustomFunctions 1.1 only</span></span> | <span data-ttu-id="f96d8-169">customfunctions 1.1 以降</span><span class="sxs-lookup"><span data-stu-id="f96d8-169">CustomFunctions 1.1 and later</span></span> |

## <a name="see-also"></a><span data-ttu-id="f96d8-170">関連項目</span><span class="sxs-lookup"><span data-stu-id="f96d8-170">See also</span></span>

- [<span data-ttu-id="f96d8-171">既存の COM アドインと互換性のある Office アドインを作成する</span><span class="sxs-lookup"><span data-stu-id="f96d8-171">Make your Office Add-in compatible with an existing COM add-in</span></span>](../develop/make-office-add-in-compatible-with-existing-com-add-in.md)
- [<span data-ttu-id="f96d8-172">カスタム関数のベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="f96d8-172">Custom functions best practices</span></span>](custom-functions-best-practices.md)
- [<span data-ttu-id="f96d8-173">カスタム関数の変更ログ</span><span class="sxs-lookup"><span data-stu-id="f96d8-173">Custom functions changelog</span></span>](custom-functions-changelog.md)
- [<span data-ttu-id="f96d8-174">Excel カスタム関数のチュートリアル</span><span class="sxs-lookup"><span data-stu-id="f96d8-174">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)