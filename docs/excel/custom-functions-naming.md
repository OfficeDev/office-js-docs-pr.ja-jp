---
ms.date: 05/17/2020
description: Excel カスタム関数の名前の要件について説明し、一般的な名前付けの落とし穴を回避します。
title: Excel のカスタム関数の名前付けガイドライン
localization_priority: Normal
ms.openlocfilehash: ac0d824f49d359e574a0dc5caae8ef2f903dd4a1
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/08/2020
ms.locfileid: "44609290"
---
# <a name="naming-guidelines"></a><span data-ttu-id="9d841-103">名前付けのガイドライン</span><span class="sxs-lookup"><span data-stu-id="9d841-103">Naming guidelines</span></span>

<span data-ttu-id="9d841-104">カスタム関数は、 `id` `name` JSON メタデータファイルのおよびプロパティによって識別されます。</span><span class="sxs-lookup"><span data-stu-id="9d841-104">A custom function is identified by an `id` and `name` property in the JSON metadata file.</span></span>

- <span data-ttu-id="9d841-105">この関数 `id` は、JavaScript コードのカスタム関数を一意に識別するために使用されます。</span><span class="sxs-lookup"><span data-stu-id="9d841-105">The function `id` is used to uniquely identify custom functions in your JavaScript code.</span></span>
- <span data-ttu-id="9d841-106">関数 `name` は、Excel でユーザーに表示される表示名として使用されます。</span><span class="sxs-lookup"><span data-stu-id="9d841-106">The function `name` is used as the display name that appears to a user in Excel.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="9d841-107">関数は、 `name` ローカライズのためなど、関数とは異なる場合が `id` あります。</span><span class="sxs-lookup"><span data-stu-id="9d841-107">A function `name` can differ from the function `id`, such as for localization purposes.</span></span> <span data-ttu-id="9d841-108">通常、関数は `name` 、 `id` それらを区別する理由がない場合は、と同じです。</span><span class="sxs-lookup"><span data-stu-id="9d841-108">In general, a function's `name` should stay the same as the `id` if there is no reason for them to differ.</span></span>

<span data-ttu-id="9d841-109">`name` `id` いくつかの一般的な要件を共有します。</span><span class="sxs-lookup"><span data-stu-id="9d841-109">A function's `name` and `id` share some common requirements:</span></span>

- <span data-ttu-id="9d841-110">関数では `id` 、a ~ Z の文字を使用することはできません。数字 0 ~ 9、アンダースコア、ピリオド。</span><span class="sxs-lookup"><span data-stu-id="9d841-110">A function's `id` may only use characters A through Z, numbers zero through nine, underscores, and periods.</span></span>

- <span data-ttu-id="9d841-111">関数では、 `name` Unicode のアルファベット文字、アンダースコア、ピリオドを使用できます。</span><span class="sxs-lookup"><span data-stu-id="9d841-111">A function's `name` may use any Unicode alphabetic characters, underscores, and periods.</span></span>

- <span data-ttu-id="9d841-112">どちらの関数も、 `name` `id` 文字で始まる必要があり、最小で3文字の制限があります。</span><span class="sxs-lookup"><span data-stu-id="9d841-112">Both function `name` and `id` must start with a letter and have a minimum limit of three characters.</span></span>

<span data-ttu-id="9d841-113">Excel は、組み込み関数名 (など) に大文字を使用 `SUM` します。</span><span class="sxs-lookup"><span data-stu-id="9d841-113">Excel uses uppercase letters for built-in function names (such as `SUM`).</span></span> <span data-ttu-id="9d841-114">カスタム関数の大文字を使用し `name` 、 `id` ベストプラクティスとして使用します。</span><span class="sxs-lookup"><span data-stu-id="9d841-114">Use uppercase letters for your custom function's `name` and `id` as a best practice.</span></span>

<span data-ttu-id="9d841-115">関数は `name` 、次のようなものである必要があります。</span><span class="sxs-lookup"><span data-stu-id="9d841-115">A function's `name` shouldn't be the same as:</span></span>

- <span data-ttu-id="9d841-116">A1 から XFD1048576 のセル、または R1C1 から R1048576C16384 までのセル。</span><span class="sxs-lookup"><span data-stu-id="9d841-116">Any cells between A1 to XFD1048576 or any cells between R1C1 to R1048576C16384.</span></span>

- <span data-ttu-id="9d841-117">任意の Excel 4.0 マクロ関数 ( `RUN` 、など `ECHO` )。</span><span class="sxs-lookup"><span data-stu-id="9d841-117">Any Excel 4.0 Macro Function (such as `RUN`, `ECHO`).</span></span>  <span data-ttu-id="9d841-118">これらの関数の完全な一覧については、「 [Excel マクロ関数リファレンスドキュメント](https://d13ot9o61jdzpp.cloudfront.net/files/Excel%204.0%20Macro%20Functions%20Reference.pdf)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="9d841-118">For a full list of these functions, see [this Excel Macro Functions Reference document](https://d13ot9o61jdzpp.cloudfront.net/files/Excel%204.0%20Macro%20Functions%20Reference.pdf).</span></span>

## <a name="naming-conflicts"></a><span data-ttu-id="9d841-119">名前付けの競合</span><span class="sxs-lookup"><span data-stu-id="9d841-119">Naming conflicts</span></span>

<span data-ttu-id="9d841-120">関数 `name` が `name` 既に存在するアドインの関数と同じ場合は、 **#REF!**</span><span class="sxs-lookup"><span data-stu-id="9d841-120">If your function `name` is the same as a function `name` in an add-in that already exists, the **#REF!**</span></span> <span data-ttu-id="9d841-121">エラーがブックに表示されます。</span><span class="sxs-lookup"><span data-stu-id="9d841-121">error will appear in your workbook.</span></span>

<span data-ttu-id="9d841-122">名前付けの競合を修正するには、アドインでを変更して、関数を再度実行し `name` ます。</span><span class="sxs-lookup"><span data-stu-id="9d841-122">To fix a naming conflict, change the `name` in your add-in and try the function again.</span></span> <span data-ttu-id="9d841-123">競合する名前を使用してアドインをアンインストールすることもできます。</span><span class="sxs-lookup"><span data-stu-id="9d841-123">You can also uninstall the add-in with the conflicting name.</span></span> <span data-ttu-id="9d841-124">または、別の環境でアドインをテストしている場合は、別の名前空間を使用して、関数を区別します (など `NAMESPACE_NAMEOFFUNCTION` )。</span><span class="sxs-lookup"><span data-stu-id="9d841-124">Or, if you're testing your add-in in different environments, try using a different namespace to differentiate your function (such as `NAMESPACE_NAMEOFFUNCTION`).</span></span>

## <a name="best-practices"></a><span data-ttu-id="9d841-125">ベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="9d841-125">Best practices</span></span>

- <span data-ttu-id="9d841-126">同じまたは似た名前を持つ複数の関数を作成するのではなく、複数の引数を関数に追加することを検討してください。</span><span class="sxs-lookup"><span data-stu-id="9d841-126">Consider adding multiple arguments to a function rather than creating multiple functions with the same or similar names.</span></span>
- <span data-ttu-id="9d841-127">関数名にあいまいな略語を含めないでください。</span><span class="sxs-lookup"><span data-stu-id="9d841-127">Avoid ambiguous abbreviations in function names.</span></span> <span data-ttu-id="9d841-128">わかりやすくすることが重要です。</span><span class="sxs-lookup"><span data-stu-id="9d841-128">Clarity is more important than brevity.</span></span> <span data-ttu-id="9d841-129">ではなく、という名前を選択し `=INCREASETIME` `=INC` ます。</span><span class="sxs-lookup"><span data-stu-id="9d841-129">Choose a name like `=INCREASETIME` rather than `=INC`.</span></span>
- <span data-ttu-id="9d841-130">関数名は、関数のアクション (ZIPCODE ではなく = GETZIPCODE など) を示す必要があります。</span><span class="sxs-lookup"><span data-stu-id="9d841-130">Function names should indicate the action of the function, such as =GETZIPCODE instead of ZIPCODE.</span></span>
- <span data-ttu-id="9d841-131">類似のアクションを実行する関数に対して同じ動詞を一貫して使用します。</span><span class="sxs-lookup"><span data-stu-id="9d841-131">Consistently use the same verbs for functions which perform similar actions.</span></span> <span data-ttu-id="9d841-132">たとえば、とで `=DELETEZIPCODE` はなくを使用し `=DELETEADDRESS` `=DELETEZIPCODE` `=REMOVEADDRESS` ます。</span><span class="sxs-lookup"><span data-stu-id="9d841-132">For example, use `=DELETEZIPCODE` and `=DELETEADDRESS`, rather than `=DELETEZIPCODE` and `=REMOVEADDRESS`.</span></span>
- <span data-ttu-id="9d841-133">ストリーミング関数の名前を指定するときは、その効果にメモを追加するか、関数の `STREAM` 名前の末尾に追加することを検討してください。</span><span class="sxs-lookup"><span data-stu-id="9d841-133">When naming a streaming function, consider adding a note to that effect in the description of the function or adding `STREAM` to the end of the function's name.</span></span>

[!include[manifest guidance](../includes/manifest-guidance.md)]

## <a name="localizing-function-names"></a><span data-ttu-id="9d841-134">関数名のローカライズ</span><span class="sxs-lookup"><span data-stu-id="9d841-134">Localizing function names</span></span>

<span data-ttu-id="9d841-135">個別の JSON ファイルを使用し、アドインのマニフェストファイルで値をオーバーライドすることにより、異なる言語の関数名をローカライズできます。</span><span class="sxs-lookup"><span data-stu-id="9d841-135">You can localize your function names for different languages using separate JSON files and override values in your add-in's manifest file.</span></span> <span data-ttu-id="9d841-136">`id`ローカライズされた関数と競合する可能性があるため、関数に、または `name` 別の言語の組み込みの Excel 関数を付与しないでください。</span><span class="sxs-lookup"><span data-stu-id="9d841-136">Avoid giving your functions an `id` or `name` that is a built-in Excel function in another language as this could conflict with localized functions.</span></span>

<span data-ttu-id="9d841-137">ローカライズの詳細については、「[カスタム関数をローカライズ](custom-functions-localize.md)する」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="9d841-137">For full information on localizing, see [Localize custom functions](custom-functions-localize.md)</span></span>

## <a name="next-steps"></a><span data-ttu-id="9d841-138">次の手順</span><span class="sxs-lookup"><span data-stu-id="9d841-138">Next steps</span></span>
<span data-ttu-id="9d841-139">[エラー処理のベストプラクティス](custom-functions-errors.md)について説明します。</span><span class="sxs-lookup"><span data-stu-id="9d841-139">Learn about [error handling best practices](custom-functions-errors.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="9d841-140">関連項目</span><span class="sxs-lookup"><span data-stu-id="9d841-140">See also</span></span>

* [<span data-ttu-id="9d841-141">カスタム関数のメタデータ</span><span class="sxs-lookup"><span data-stu-id="9d841-141">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="9d841-142">Excel カスタム関数のチュートリアル</span><span class="sxs-lookup"><span data-stu-id="9d841-142">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
