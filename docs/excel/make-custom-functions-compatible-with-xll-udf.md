---
title: XLL ユーザー定義関数を使用してカスタム関数を拡張する
description: カスタム関数と同等の機能を持つ Excel XLL ユーザー定義関数との互換性を有効にする
ms.date: 07/31/2019
localization_priority: Normal
ms.openlocfilehash: 7ec853e5b4d03267e1c9d33d2df8a79d86860095
ms.sourcegitcommit: c8914ce0f48a0c19bbfc3276a80d090bb7ce68e1
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/26/2019
ms.locfileid: "37235303"
---
# <a name="extend-custom-functions-with-xll-user-defined-functions"></a><span data-ttu-id="bd7d1-103">XLL ユーザー定義関数を使用してカスタム関数を拡張する</span><span class="sxs-lookup"><span data-stu-id="bd7d1-103">Extend custom functions with XLL user-defined functions</span></span>

<span data-ttu-id="bd7d1-104">既存の Excel XLLs がある場合は、Excel アドインで同等のカスタム関数を作成して、online や macOS などの他のプラットフォームにソリューション機能を拡張することができます。</span><span class="sxs-lookup"><span data-stu-id="bd7d1-104">If you have existing Excel XLLs, you can build equivalent custom functions in an Excel add-in to extend your solution features to other platforms such as online or macOS.</span></span> <span data-ttu-id="bd7d1-105">ただし、Excel アドインには、xll で利用可能なすべての機能が含まれているわけではありません。</span><span class="sxs-lookup"><span data-stu-id="bd7d1-105">However, Excel add-ins don't have all of the functionality available in XLLs.</span></span> <span data-ttu-id="bd7d1-106">ソリューションで使用されている機能によっては、XLL の方が excel の excel アドインカスタム関数よりも優れた操作を提供することがあります。</span><span class="sxs-lookup"><span data-stu-id="bd7d1-106">Depending on the functionality your solution uses, the XLL may provide a better experience than the Excel add-in custom functions in Excel on Windows.</span></span>

> [!NOTE]
> <span data-ttu-id="bd7d1-107">COM アドインと XLL の UDF の互換性は、Office 365 サブスクリプションに接続している場合、次のプラットフォームでサポートされています。</span><span class="sxs-lookup"><span data-stu-id="bd7d1-107">COM add-in and XLL UDF compatibility is supported by the following platforms, when connected to an Office 365 subscription:</span></span>
> - <span data-ttu-id="bd7d1-108">Excel on the web</span><span class="sxs-lookup"><span data-stu-id="bd7d1-108">Excel on the web</span></span>
> - <span data-ttu-id="bd7d1-109">Windows 版 Excel (バージョン1904以降)</span><span class="sxs-lookup"><span data-stu-id="bd7d1-109">Excel on Windows (version 1904 or later)</span></span>
> - <span data-ttu-id="bd7d1-110">Excel on Mac (バージョン13.329 以降)</span><span class="sxs-lookup"><span data-stu-id="bd7d1-110">Excel on Mac (version 13.329 or later)</span></span>
> 
> <span data-ttu-id="bd7d1-111">Web 上の Excel で COM アドインと XLL UDF との互換性を使用するには、Office 365 サブスクリプションまたは[Microsoft アカウント](https://account.microsoft.com/account)のいずれかを使用してログインします。</span><span class="sxs-lookup"><span data-stu-id="bd7d1-111">To use COM add-in and XLL UDF compatibility within Excel on the web, login by using either your Office 365 subscription or a [Microsoft account](https://account.microsoft.com/account).</span></span> <span data-ttu-id="bd7d1-112">Office 365 サブスクリプションをまだお持ちでない場合は、[Office 365 Developer Program](https://developer.microsoft.com/office/dev-program) に参加することで入手できます。</span><span class="sxs-lookup"><span data-stu-id="bd7d1-112">If you don't already have an Office 365 subscription, you can get one by joining the [Office 365 Developer Program](https://developer.microsoft.com/office/dev-program).</span></span>

## <a name="specify-equivalent-xll-in-the-manifest"></a><span data-ttu-id="bd7d1-113">マニフェストで同等の XLL を指定する</span><span class="sxs-lookup"><span data-stu-id="bd7d1-113">Specify equivalent XLL in the manifest</span></span>

<span data-ttu-id="bd7d1-114">既存の XLL との互換性を有効にするには、Excel アドインのマニフェストで同等の XLL を識別します。</span><span class="sxs-lookup"><span data-stu-id="bd7d1-114">To enable compatibility with an existing XLL, identify the equivalent XLL in the manifest of your Excel add-in.</span></span> <span data-ttu-id="bd7d1-115">Excel では、Windows での実行時に Excel アドインカスタム関数の代わりに XLL 関数が使用されます。</span><span class="sxs-lookup"><span data-stu-id="bd7d1-115">Then Excel will use the XLL's functions instead of your Excel add-in custom functions when running on Windows.</span></span>

<span data-ttu-id="bd7d1-116">カスタム関数に対応する XLL を設定するには、 `FileName` xll のを指定します。</span><span class="sxs-lookup"><span data-stu-id="bd7d1-116">To set the equivalent XLL for your custom functions, specify the `FileName` of the XLL.</span></span> <span data-ttu-id="bd7d1-117">ユーザーが XLL から関数を含むブックを開くと、Excel は関数を互換性のある関数に変換します。</span><span class="sxs-lookup"><span data-stu-id="bd7d1-117">When the user opens a workbook with functions from the XLL, Excel converts the functions to compatible functions.</span></span> <span data-ttu-id="bd7d1-118">ブックは、Windows の Excel で開いたときに XLL を使用し、オンラインまたは macOS を開いたときに Excel アドインのカスタム関数を使用します。</span><span class="sxs-lookup"><span data-stu-id="bd7d1-118">The workbook then uses the XLL when opened in Excel on Windows, and it will use custom functions from your Excel add-in when opened online or on macOS.</span></span>

<span data-ttu-id="bd7d1-119">次の例は、COM アドインと XLL の両方を同等として指定する方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="bd7d1-119">The following example shows how to specify both a COM add-in and an XLL as equivalent.</span></span> <span data-ttu-id="bd7d1-120">多くの場合、この例は完全にコンテキストで指定します。</span><span class="sxs-lookup"><span data-stu-id="bd7d1-120">Often you will specify both so for completeness this example shows both in context.</span></span> <span data-ttu-id="bd7d1-121">これらは、 `FileName`それぞれに`ProgId`よって識別されます。</span><span class="sxs-lookup"><span data-stu-id="bd7d1-121">They are identified by their `ProgId` and `FileName` respectively.</span></span> <span data-ttu-id="bd7d1-122">要素`EquivalentAddins`は、終了`VersionOverrides`タグの直前に配置する必要があります。</span><span class="sxs-lookup"><span data-stu-id="bd7d1-122">The `EquivalentAddins` element must be positioned immediately before the closing `VersionOverrides` tag.</span></span> <span data-ttu-id="bd7d1-123">COM アドインの互換性の詳細については、「[既存の com アドインと互換性のある Excel アドインを作成](../develop/make-office-add-in-compatible-with-existing-com-add-in.md)する」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="bd7d1-123">For more information on COM add-in compatibility, see [Make your Excel add-in compatible with an existing COM add-in](../develop/make-office-add-in-compatible-with-existing-com-add-in.md).</span></span>

```xml
<VersionOverrides>
  ...
  <EquivalentAddins>
    <EquivalentAddin>
      <ProgId>ContosoCOMAddin</ProgId>
      <Type>COM</Type>
    </EquivalentAddin>

    <EquivalentAddin>
      <FileName>contosofunctions.xll</FileName>
      <Type>XLL</Type>
    </EquivalentAddin>
  <EquivalentAddins>
</VersionOverrides>
```

> [!NOTE]
> <span data-ttu-id="bd7d1-124">アドインでカスタム関数が XLL 互換に宣言されている場合、後でマニフェストを変更すると、ファイル形式が変更されるため、ユーザーのブックが破損する可能性があります。</span><span class="sxs-lookup"><span data-stu-id="bd7d1-124">If an add-in declares its custom functions to be XLL compatible, changing the manifest at a later time could break a user’s workbook because it will change the file format.</span></span>

## <a name="excel-add-in-updates"></a><span data-ttu-id="bd7d1-125">Excel アドインの更新プログラム</span><span class="sxs-lookup"><span data-stu-id="bd7d1-125">Excel add-in updates</span></span>

<span data-ttu-id="bd7d1-126">Excel アドインに対して同等の XLL を指定すると、excel アドインの更新プログラムの処理は中止されます。</span><span class="sxs-lookup"><span data-stu-id="bd7d1-126">Once you specify an equivalent XLL for your Excel add-in, Excel stops processing updates for your Excel add-in.</span></span> <span data-ttu-id="bd7d1-127">ユーザーは、Excel アドインの最新の更新プログラムを取得するために、XLL をアンインストールする必要があります。</span><span class="sxs-lookup"><span data-stu-id="bd7d1-127">The user must uninstall the XLL in order to get the latest updates for the Excel add-in.</span></span>

## <a name="custom-function-behavior-for-xll-compatible-functions"></a><span data-ttu-id="bd7d1-128">XLL 互換関数のカスタム関数の動作</span><span class="sxs-lookup"><span data-stu-id="bd7d1-128">Custom function behavior for XLL compatible functions</span></span>

<span data-ttu-id="bd7d1-129">同じアドインが含まれている XLL 関数を含むスプレッドシートが開かれると、xll 関数は、XLL 互換のカスタム関数に変換されます。</span><span class="sxs-lookup"><span data-stu-id="bd7d1-129">When a spreadsheet is opened that contains XLL functions for which there is also an equivalent add-in, the XLL's functions are converted to XLL compatible custom functions.</span></span> <span data-ttu-id="bd7d1-130">次の保存時に、これらのファイルは互換モードでファイルに書き込まれます。これにより、(他のプラットフォームでの場合) XLL と Excel アドインの両方のカスタム機能を使用できるようになります。</span><span class="sxs-lookup"><span data-stu-id="bd7d1-130">On the next save, they are written to the file in a compatible mode so that they work with both the XLL and Excel add-in custom functions (when on other platforms).</span></span>

<span data-ttu-id="bd7d1-131">次の表は、XLL ユーザー定義関数、XLL 互換カスタム関数、および Excel アドインカスタム関数の機能を比較しています。</span><span class="sxs-lookup"><span data-stu-id="bd7d1-131">The following table compares features across XLL user-defined functions, XLL compatible custom functions, and Excel add-in custom functions.</span></span>

|         |<span data-ttu-id="bd7d1-132">XLL のユーザー定義関数</span><span class="sxs-lookup"><span data-stu-id="bd7d1-132">XLL user-defined function</span></span> |<span data-ttu-id="bd7d1-133">XLL 互換のカスタム関数</span><span class="sxs-lookup"><span data-stu-id="bd7d1-133">XLL compatible custom functions</span></span> |<span data-ttu-id="bd7d1-134">Excel アドインのカスタム関数</span><span class="sxs-lookup"><span data-stu-id="bd7d1-134">Excel add-in custom function</span></span> |
|---------|---------|---------|---------|
| <span data-ttu-id="bd7d1-135">サポートされるプラットフォーム</span><span class="sxs-lookup"><span data-stu-id="bd7d1-135">Supported platforms</span></span> | <span data-ttu-id="bd7d1-136">Windows</span><span class="sxs-lookup"><span data-stu-id="bd7d1-136">Windows</span></span> | <span data-ttu-id="bd7d1-137">Windows、macOS、Excel on the web</span><span class="sxs-lookup"><span data-stu-id="bd7d1-137">Windows, macOS, Excel on the web</span></span> | <span data-ttu-id="bd7d1-138">Windows、macOS、Excel on the web</span><span class="sxs-lookup"><span data-stu-id="bd7d1-138">Windows, macOS, Excel on the web</span></span> |
| <span data-ttu-id="bd7d1-139">サポートされるファイル形式</span><span class="sxs-lookup"><span data-stu-id="bd7d1-139">Supported file formats</span></span> | <span data-ttu-id="bd7d1-140">.XLSX、.XLSB、.XLSM、XLS</span><span class="sxs-lookup"><span data-stu-id="bd7d1-140">XLSX, XLSB, XLSM, XLS</span></span> | <span data-ttu-id="bd7d1-141">.XLSX、.XLSB、.XLSM</span><span class="sxs-lookup"><span data-stu-id="bd7d1-141">XLSX, XLSB, XLSM</span></span> | <span data-ttu-id="bd7d1-142">.XLSX、.XLSB、.XLSM</span><span class="sxs-lookup"><span data-stu-id="bd7d1-142">XLSX, XLSB, XLSM</span></span> |
| <span data-ttu-id="bd7d1-143">数式オートコンプリート</span><span class="sxs-lookup"><span data-stu-id="bd7d1-143">Formula autocomplete</span></span> | <span data-ttu-id="bd7d1-144">いいえ</span><span class="sxs-lookup"><span data-stu-id="bd7d1-144">No</span></span> | <span data-ttu-id="bd7d1-145">はい</span><span class="sxs-lookup"><span data-stu-id="bd7d1-145">Yes</span></span> | <span data-ttu-id="bd7d1-146">はい</span><span class="sxs-lookup"><span data-stu-id="bd7d1-146">Yes</span></span> |
| <span data-ttu-id="bd7d1-147">ストリーミング</span><span class="sxs-lookup"><span data-stu-id="bd7d1-147">Streaming</span></span> | <span data-ttu-id="bd7d1-148">XlfRTD および XLL コールバックを使用して可能。</span><span class="sxs-lookup"><span data-stu-id="bd7d1-148">Possible via xlfRTD and XLL callback.</span></span> | <span data-ttu-id="bd7d1-149">いいえ</span><span class="sxs-lookup"><span data-stu-id="bd7d1-149">No</span></span> | <span data-ttu-id="bd7d1-150">はい</span><span class="sxs-lookup"><span data-stu-id="bd7d1-150">Yes</span></span> |
| <span data-ttu-id="bd7d1-151">関数のローカライズ</span><span class="sxs-lookup"><span data-stu-id="bd7d1-151">Localization of functions</span></span> | <span data-ttu-id="bd7d1-152">いいえ</span><span class="sxs-lookup"><span data-stu-id="bd7d1-152">No</span></span> | <span data-ttu-id="bd7d1-153">いいえ。</span><span class="sxs-lookup"><span data-stu-id="bd7d1-153">No.</span></span> <span data-ttu-id="bd7d1-154">名前と ID は、既存の XLL 関数と一致している必要があります。</span><span class="sxs-lookup"><span data-stu-id="bd7d1-154">The Name and ID must match the existing XLL's functions.</span></span> | <span data-ttu-id="bd7d1-155">はい</span><span class="sxs-lookup"><span data-stu-id="bd7d1-155">Yes</span></span> |
| <span data-ttu-id="bd7d1-156">揮発性関数</span><span class="sxs-lookup"><span data-stu-id="bd7d1-156">Volatile functions</span></span> | <span data-ttu-id="bd7d1-157">はい</span><span class="sxs-lookup"><span data-stu-id="bd7d1-157">Yes</span></span> | <span data-ttu-id="bd7d1-158">はい</span><span class="sxs-lookup"><span data-stu-id="bd7d1-158">Yes</span></span> | <span data-ttu-id="bd7d1-159">はい</span><span class="sxs-lookup"><span data-stu-id="bd7d1-159">Yes</span></span> |
| <span data-ttu-id="bd7d1-160">マルチスレッドの再計算のサポート</span><span class="sxs-lookup"><span data-stu-id="bd7d1-160">Multi-threaded recalculation support</span></span> | <span data-ttu-id="bd7d1-161">はい</span><span class="sxs-lookup"><span data-stu-id="bd7d1-161">Yes</span></span> | <span data-ttu-id="bd7d1-162">はい</span><span class="sxs-lookup"><span data-stu-id="bd7d1-162">Yes</span></span> | <span data-ttu-id="bd7d1-163">はい</span><span class="sxs-lookup"><span data-stu-id="bd7d1-163">Yes</span></span> |
| <span data-ttu-id="bd7d1-164">計算動作</span><span class="sxs-lookup"><span data-stu-id="bd7d1-164">Calculation behavior</span></span> | <span data-ttu-id="bd7d1-165">UI がありません。</span><span class="sxs-lookup"><span data-stu-id="bd7d1-165">No UI.</span></span> <span data-ttu-id="bd7d1-166">計算中に Excel が応答しなくなることがあります。</span><span class="sxs-lookup"><span data-stu-id="bd7d1-166">Excel can be unresponsive during calculation.</span></span> | <span data-ttu-id="bd7d1-167">ユーザーには #BUSY が表示されます。</span><span class="sxs-lookup"><span data-stu-id="bd7d1-167">Users will see #BUSY!</span></span> <span data-ttu-id="bd7d1-168">を返します。</span><span class="sxs-lookup"><span data-stu-id="bd7d1-168">until a result is returned.</span></span> | <span data-ttu-id="bd7d1-169">ユーザーには #BUSY が表示されます。</span><span class="sxs-lookup"><span data-stu-id="bd7d1-169">Users will see #BUSY!</span></span> <span data-ttu-id="bd7d1-170">を返します。</span><span class="sxs-lookup"><span data-stu-id="bd7d1-170">until a result is returned.</span></span> |
| <span data-ttu-id="bd7d1-171">要件セット</span><span class="sxs-lookup"><span data-stu-id="bd7d1-171">Requirement sets</span></span> | <span data-ttu-id="bd7d1-172">なし。</span><span class="sxs-lookup"><span data-stu-id="bd7d1-172">N/A</span></span> | <span data-ttu-id="bd7d1-173">CustomFunctions 1.1 以降</span><span class="sxs-lookup"><span data-stu-id="bd7d1-173">CustomFunctions 1.1 and later</span></span> | <span data-ttu-id="bd7d1-174">CustomFunctions 1.1 以降</span><span class="sxs-lookup"><span data-stu-id="bd7d1-174">CustomFunctions 1.1 and later</span></span> |

## <a name="see-also"></a><span data-ttu-id="bd7d1-175">関連項目</span><span class="sxs-lookup"><span data-stu-id="bd7d1-175">See also</span></span>

- [<span data-ttu-id="bd7d1-176">既存の COM アドインと互換性のある Excel アドインを作成する</span><span class="sxs-lookup"><span data-stu-id="bd7d1-176">Make your Excel add-in compatible with an existing COM add-in</span></span>](../develop/make-office-add-in-compatible-with-existing-com-add-in.md)
- [<span data-ttu-id="bd7d1-177">チュートリアル: Excel でカスタム関数を作成します。</span><span class="sxs-lookup"><span data-stu-id="bd7d1-177">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
