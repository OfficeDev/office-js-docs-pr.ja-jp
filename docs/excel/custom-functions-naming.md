---
ms.date: 12/28/2019
description: Excel カスタム関数の名前の要件について説明し、一般的な名前付けの落とし穴を回避します。
title: Excel のカスタム関数の名前付けガイドライン
localization_priority: Normal
ms.openlocfilehash: 81ce0e1a1d510fd9558a3e57273903382326ad55
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/17/2020
ms.locfileid: "42719477"
---
# <a name="naming-guidelines"></a><span data-ttu-id="2b2c5-103">名前付けのガイドライン</span><span class="sxs-lookup"><span data-stu-id="2b2c5-103">Naming guidelines</span></span>

<span data-ttu-id="2b2c5-104">カスタム関数は、JSON メタデータ`id`ファイル`name`のおよびプロパティによって識別されます。</span><span class="sxs-lookup"><span data-stu-id="2b2c5-104">A custom function is identified by an `id` and `name` property in the JSON metadata file.</span></span>

- <span data-ttu-id="2b2c5-105">この関数`id`は、JavaScript コードのカスタム関数を一意に識別するために使用されます。</span><span class="sxs-lookup"><span data-stu-id="2b2c5-105">The function `id` is used to uniquely identify custom functions in your JavaScript code.</span></span>
- <span data-ttu-id="2b2c5-106">関数`name`は、Excel でユーザーに表示される表示名として使用されます。</span><span class="sxs-lookup"><span data-stu-id="2b2c5-106">The function `name` is used as the display name that appears to a user in Excel.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="2b2c5-107">関数`name`は、ローカライズのためなど`id`、関数とは異なる場合があります。</span><span class="sxs-lookup"><span data-stu-id="2b2c5-107">A function `name` can differ from the function `id`, such as for localization purposes.</span></span> <span data-ttu-id="2b2c5-108">一般的に、関数の`name`違いがない場合は、 `id`関数はと同じにしておく必要があります。</span><span class="sxs-lookup"><span data-stu-id="2b2c5-108">In general, a function's `name` should stay the same as the `id` if there is no compelling reason for them to differ.</span></span>

<span data-ttu-id="2b2c5-109">いくつかの`name`一般的`id`な要件を共有します。</span><span class="sxs-lookup"><span data-stu-id="2b2c5-109">A function's `name` and `id` share some common requirements:</span></span>

- <span data-ttu-id="2b2c5-110">関数では`id` 、a ~ Z の文字を使用することはできません。数字 0 ~ 9、アンダースコア、ピリオド。</span><span class="sxs-lookup"><span data-stu-id="2b2c5-110">A function's `id` may only use characters A through Z, numbers zero through nine, underscores, and periods.</span></span>

- <span data-ttu-id="2b2c5-111">関数では`name` 、Unicode のアルファベット文字、アンダースコア、ピリオドを使用できます。</span><span class="sxs-lookup"><span data-stu-id="2b2c5-111">A function's `name` may use any Unicode alphabetic characters, underscores, and periods.</span></span>

- <span data-ttu-id="2b2c5-112">どちらの`name`関数`id`も、文字で始まる必要があり、最小で3文字の制限があります。</span><span class="sxs-lookup"><span data-stu-id="2b2c5-112">Both function `name` and `id` must start with a letter and have a minimum limit of three characters.</span></span>

<span data-ttu-id="2b2c5-113">Excel は、組み込み関数名 (など`SUM`) に大文字を使用します。</span><span class="sxs-lookup"><span data-stu-id="2b2c5-113">Excel uses uppercase letters for built-in function names (such as `SUM`).</span></span> <span data-ttu-id="2b2c5-114">そのため、カスタム関数`name`に大文字を使用し、 `id`ベストプラクティスとして使用することを検討してください。</span><span class="sxs-lookup"><span data-stu-id="2b2c5-114">Therefore, consider using uppercase letters for your custom function's `name` and `id` as a best practice.</span></span>

<span data-ttu-id="2b2c5-115">関数`name`には、次のような名前を付けることはできません。</span><span class="sxs-lookup"><span data-stu-id="2b2c5-115">A function's `name` shouldn't be named the same as:</span></span>

- <span data-ttu-id="2b2c5-116">A1 から XFD1048576 のセル、または R1C1 から R1048576C16384 までのセル。</span><span class="sxs-lookup"><span data-stu-id="2b2c5-116">Any cells between A1 to XFD1048576 or any cells between R1C1 to R1048576C16384.</span></span>

- <span data-ttu-id="2b2c5-117">任意の Excel 4.0 マクロ関数 ( `RUN`、 `ECHO`など)。</span><span class="sxs-lookup"><span data-stu-id="2b2c5-117">Any Excel 4.0 Macro Function (such as `RUN`, `ECHO`).</span></span>  <span data-ttu-id="2b2c5-118">これらの関数の完全な一覧については、「 [Excel マクロ関数リファレンスドキュメント](https://d13ot9o61jdzpp.cloudfront.net/files/Excel%204.0%20Macro%20Functions%20Reference.pdf)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="2b2c5-118">For a full list of these functions, see [this Excel Macro Functions Reference document](https://d13ot9o61jdzpp.cloudfront.net/files/Excel%204.0%20Macro%20Functions%20Reference.pdf).</span></span>

## <a name="naming-conflicts"></a><span data-ttu-id="2b2c5-119">名前付けの競合</span><span class="sxs-lookup"><span data-stu-id="2b2c5-119">Naming conflicts</span></span>

<span data-ttu-id="2b2c5-120">関数`name`が既に存在するアドインの関数`name`と同じ場合は、 **#REF!**</span><span class="sxs-lookup"><span data-stu-id="2b2c5-120">If your function `name` is the same as a function `name` in an add-in that already exists, the **#REF!**</span></span> <span data-ttu-id="2b2c5-121">エラーがブックに表示されます。</span><span class="sxs-lookup"><span data-stu-id="2b2c5-121">error will appear in your workbook.</span></span>

<span data-ttu-id="2b2c5-122">名前付けの競合を修正するに`name`は、アドインでを変更して、関数を再度実行します。</span><span class="sxs-lookup"><span data-stu-id="2b2c5-122">To fix a naming conflict, change the `name` in your add-in and try the function again.</span></span> <span data-ttu-id="2b2c5-123">競合する名前を使用してアドインをアンインストールすることもできます。</span><span class="sxs-lookup"><span data-stu-id="2b2c5-123">You can also uninstall the add-in with the conflicting name.</span></span> <span data-ttu-id="2b2c5-124">または、別の環境でアドインをテストしている場合は、別の名前空間を使用して、関数`NAMESPACE_NAMEOFFUNCTION`を区別します (など)。</span><span class="sxs-lookup"><span data-stu-id="2b2c5-124">Or, if you're testing your add-in in different environments, try using a different namespace to differentiate your function (such as `NAMESPACE_NAMEOFFUNCTION`).</span></span>

## <a name="best-practices"></a><span data-ttu-id="2b2c5-125">ベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="2b2c5-125">Best practices</span></span>

- <span data-ttu-id="2b2c5-126">同じまたは似た名前を持つ複数の関数を作成するのではなく、複数の引数を関数に追加することを検討してください。</span><span class="sxs-lookup"><span data-stu-id="2b2c5-126">Consider adding multiple arguments to a function rather than creating multiple functions with the same or similar names.</span></span>
- <span data-ttu-id="2b2c5-127">関数名は、ではなく、関数のアクションを`=GETZIPCODE`示して`ZIPCODE`いなければなりません。</span><span class="sxs-lookup"><span data-stu-id="2b2c5-127">Function names should indicate the action of the function, such as `=GETZIPCODE` instead of `ZIPCODE`.</span></span>
- <span data-ttu-id="2b2c5-128">関数名にあいまいな略語を含めないでください。</span><span class="sxs-lookup"><span data-stu-id="2b2c5-128">Avoid ambiguous abbreviations in function names.</span></span> <span data-ttu-id="2b2c5-129">わかりやすくすることが重要です。</span><span class="sxs-lookup"><span data-stu-id="2b2c5-129">Clarity is more important than brevity.</span></span> <span data-ttu-id="2b2c5-130">ではなく、 `=INCREASETIME`という`=INC`名前を選択します。</span><span class="sxs-lookup"><span data-stu-id="2b2c5-130">Choose a name like `=INCREASETIME` rather than `=INC`.</span></span>
- <span data-ttu-id="2b2c5-131">類似のアクションを実行する関数に対して同じ動詞を一貫して使用します。</span><span class="sxs-lookup"><span data-stu-id="2b2c5-131">Consistently use the same verbs for functions which perform similar actions.</span></span> <span data-ttu-id="2b2c5-132">たとえば、とで`=DELETEZIPCODE`は`=DELETEADDRESS`なく`=DELETEZIPCODE`を使用し`=REMOVEADDRESS`ます。</span><span class="sxs-lookup"><span data-stu-id="2b2c5-132">For example, use `=DELETEZIPCODE` and `=DELETEADDRESS`, rather than `=DELETEZIPCODE` and `=REMOVEADDRESS`.</span></span>
- <span data-ttu-id="2b2c5-133">ストリーミング関数の名前を指定するときは、その効果にメモを追加するか、関数の`STREAM`名前の末尾に追加することを検討してください。</span><span class="sxs-lookup"><span data-stu-id="2b2c5-133">When naming a streaming function, consider adding a note to that effect in the description of the function or adding `STREAM` to the end of the function's name.</span></span>

[!include[manifest guidance](../includes/manifest-guidance.md)]

## <a name="localizing-function-names"></a><span data-ttu-id="2b2c5-134">関数名のローカライズ</span><span class="sxs-lookup"><span data-stu-id="2b2c5-134">Localizing function names</span></span>

<span data-ttu-id="2b2c5-135">個別の JSON ファイルを使用し、アドインのマニフェストファイルで値をオーバーライドすることにより、異なる言語の関数名をローカライズできます。</span><span class="sxs-lookup"><span data-stu-id="2b2c5-135">You can localize your function names for different languages using separate JSON files and override values in your add-in's manifest file.</span></span> <span data-ttu-id="2b2c5-136">これはローカライズされた関数と競合する`id`可能性`name`があるため、ベストプラクティスとして、関数または組み込みの Excel 関数を別の言語で提供しないようにします。</span><span class="sxs-lookup"><span data-stu-id="2b2c5-136">As a best practice, avoid giving your functions an `id` or `name` that is a built-in Excel function in another language as this could conflict with localized functions.</span></span>

<span data-ttu-id="2b2c5-137">ローカライズの詳細については、「[カスタム関数をローカライズ](custom-functions-localize.md)する」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="2b2c5-137">For full information on localizing, see [Localize custom functions](custom-functions-localize.md)</span></span>

## <a name="next-steps"></a><span data-ttu-id="2b2c5-138">次の手順</span><span class="sxs-lookup"><span data-stu-id="2b2c5-138">Next steps</span></span>
<span data-ttu-id="2b2c5-139">[エラー処理のベストプラクティス](custom-functions-errors.md)について説明します。</span><span class="sxs-lookup"><span data-stu-id="2b2c5-139">Learn about [error handling best practices](custom-functions-errors.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="2b2c5-140">関連項目</span><span class="sxs-lookup"><span data-stu-id="2b2c5-140">See also</span></span>

* [<span data-ttu-id="2b2c5-141">カスタム関数のメタデータ</span><span class="sxs-lookup"><span data-stu-id="2b2c5-141">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="2b2c5-142">Excel カスタム関数のチュートリアル</span><span class="sxs-lookup"><span data-stu-id="2b2c5-142">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
* [<span data-ttu-id="2b2c5-143">Excel カスタム関数のランタイム</span><span class="sxs-lookup"><span data-stu-id="2b2c5-143">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
