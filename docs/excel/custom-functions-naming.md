---
ms.date: 05/03/2019
description: Excel カスタム関数の名前の要件について説明し、一般的な名前付けの落とし穴を回避します。
title: Excel のカスタム関数の名前付けガイドライン
localization_priority: Normal
ms.openlocfilehash: 3abe04eebfa703666b70ecbde1c68ab0c942003c
ms.sourcegitcommit: ff73cc04e5718765fcbe74181505a974db69c3f5
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/06/2019
ms.locfileid: "33628047"
---
# <a name="naming-guidelines"></a><span data-ttu-id="8fa5d-103">名前付けのガイドライン</span><span class="sxs-lookup"><span data-stu-id="8fa5d-103">Naming guidelines</span></span>

<span data-ttu-id="8fa5d-104">カスタム関数は、JSON メタデータファイルの**id**および**name**プロパティによって識別されます。</span><span class="sxs-lookup"><span data-stu-id="8fa5d-104">A custom function is identified by an **id** and **name** property in the JSON metadata file.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

- <span data-ttu-id="8fa5d-105">この関数`id`は、JavaScript コードのカスタム関数を一意に識別するために使用されます。</span><span class="sxs-lookup"><span data-stu-id="8fa5d-105">The function `id` is used to uniquely identify custom functions in your JavaScript code.</span></span> 
- <span data-ttu-id="8fa5d-106">関数`name`は、Excel でユーザーに表示される表示名として使用されます。</span><span class="sxs-lookup"><span data-stu-id="8fa5d-106">The function `name` is used as the display name that appears to a user in Excel.</span></span> 

<span data-ttu-id="8fa5d-107">関数`name`は、ローカライズのためなど`id`、関数とは異なる場合があります。</span><span class="sxs-lookup"><span data-stu-id="8fa5d-107">A function `name` can differ from the function `id`, such as for localization purposes.</span></span> <span data-ttu-id="8fa5d-108">一般的に、関数の`name`違いがない場合は、 `id`関数はと同じにしておく必要があります。</span><span class="sxs-lookup"><span data-stu-id="8fa5d-108">In general, a function's `name` should stay the same as the `id` if there is no compelling reason for them to differ.</span></span>

<span data-ttu-id="8fa5d-109">いくつかの`name`一般的`id`な要件を共有します。</span><span class="sxs-lookup"><span data-stu-id="8fa5d-109">A function's `name` and `id` share some common requirements:</span></span>

- <span data-ttu-id="8fa5d-110">関数では`id` 、A ~ Z の文字を使用することはできません。数字 0 ~ 9、アンダースコア、ピリオド。</span><span class="sxs-lookup"><span data-stu-id="8fa5d-110">A function's `id` may only use characters A through Z, numbers zero through nine, underscores, and periods.</span></span>

- <span data-ttu-id="8fa5d-111">関数では`name` 、Unicode のアルファベット文字、アンダースコア、ピリオドを使用できます。</span><span class="sxs-lookup"><span data-stu-id="8fa5d-111">A function's `name` may use any Unicode alphabetic characters, underscores, and periods.</span></span>

- <span data-ttu-id="8fa5d-112">どちらの`name`関数`id`も、文字で始まる必要があり、最小で3文字の制限があります。</span><span class="sxs-lookup"><span data-stu-id="8fa5d-112">Both function `name` and `id` must start with a letter and have a minimum limit of three characters.</span></span>

<span data-ttu-id="8fa5d-113">Excel は、組み込み関数名 (など`SUM`) に大文字を使用します。</span><span class="sxs-lookup"><span data-stu-id="8fa5d-113">Excel uses uppercase letters for built-in function names (such as `SUM`).</span></span> <span data-ttu-id="8fa5d-114">そのため、カスタム関数`name`に大文字を使用し、 `id`ベストプラクティスとして使用することを検討してください。</span><span class="sxs-lookup"><span data-stu-id="8fa5d-114">Therefore, consider using uppercase letters for your custom function's `name` and `id` as a best practice.</span></span>

<span data-ttu-id="8fa5d-115">関数`name`には、次のような名前を付けることはできません。</span><span class="sxs-lookup"><span data-stu-id="8fa5d-115">A function's `name` shouldn't be named the same as:</span></span>

- <span data-ttu-id="8fa5d-116">A1 から XFD1048576 のセル、または R1C1 から R1048576C16384 までのセル。</span><span class="sxs-lookup"><span data-stu-id="8fa5d-116">Any cells between A1 to XFD1048576 or any cells between R1C1 to R1048576C16384.</span></span>

- <span data-ttu-id="8fa5d-117">任意の Excel 4.0 マクロ関数 ( `RUN`、 `ECHO`など)。</span><span class="sxs-lookup"><span data-stu-id="8fa5d-117">Any Excel 4.0 Macro Function (such as `RUN`, `ECHO`).</span></span>  <span data-ttu-id="8fa5d-118">これらの関数の完全な一覧については、[この記事](https://www.microsoft.com/en-us/download/details.aspx?id=1465)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="8fa5d-118">For a full list of these functions, see [this article](https://www.microsoft.com/en-us/download/details.aspx?id=1465).</span></span>

## <a name="naming-conflicts"></a><span data-ttu-id="8fa5d-119">名前付けの競合</span><span class="sxs-lookup"><span data-stu-id="8fa5d-119">Naming conflicts</span></span>

<span data-ttu-id="8fa5d-120">関数`name`が既に存在するアドインの関数`name`と同じ場合は、 **#REF!**</span><span class="sxs-lookup"><span data-stu-id="8fa5d-120">If your function `name` is the same as a function `name` in an add-in that already exists, the **#REF!**</span></span> <span data-ttu-id="8fa5d-121">エラーがブックに表示されます。</span><span class="sxs-lookup"><span data-stu-id="8fa5d-121">error will appear in your workbook.</span></span>

<span data-ttu-id="8fa5d-122">名前付けの競合を修正するに`name`は、アドインでを変更して、関数を再度実行します。</span><span class="sxs-lookup"><span data-stu-id="8fa5d-122">To fix a naming conflict, change the `name` in your add-in and try the function again.</span></span> <span data-ttu-id="8fa5d-123">競合する名前を使用してアドインをアンインストールすることもできます。</span><span class="sxs-lookup"><span data-stu-id="8fa5d-123">You can also uninstall the add-in with the conflicting name.</span></span> <span data-ttu-id="8fa5d-124">または、別の環境でアドインをテストしている場合は、別の名前空間を使用して、関数`NAMESPACE_NAMEOFFUNCTION`を区別します (など)。</span><span class="sxs-lookup"><span data-stu-id="8fa5d-124">Or, if you're testing your add-in in different environments, try using a different namespace to differentiate your function (such as `NAMESPACE_NAMEOFFUNCTION`).</span></span>

## <a name="best-practices"></a><span data-ttu-id="8fa5d-125">ベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="8fa5d-125">Best practices</span></span>

- <span data-ttu-id="8fa5d-126">同じまたは似た名前を持つ複数の関数を作成するのではなく、複数の引数を関数に追加することを検討してください。</span><span class="sxs-lookup"><span data-stu-id="8fa5d-126">Consider adding multiple arguments to a function rather than creating multiple functions with the same or similar names.</span></span>
- <span data-ttu-id="8fa5d-127">関数名は、ではなく、関数のアクションを`=GETZIPCODE`示して`ZIPCODE`いなければなりません。</span><span class="sxs-lookup"><span data-stu-id="8fa5d-127">Function names should indicate the action of the function, such as `=GETZIPCODE` instead of `ZIPCODE`.</span></span>
- <span data-ttu-id="8fa5d-128">関数名にあいまいな略語を含めないでください。</span><span class="sxs-lookup"><span data-stu-id="8fa5d-128">Avoid ambiguous abbreviations in function names.</span></span> <span data-ttu-id="8fa5d-129">わかりやすくすることが重要です。</span><span class="sxs-lookup"><span data-stu-id="8fa5d-129">Clarity is more important than brevity.</span></span> <span data-ttu-id="8fa5d-130">ではなく、 `=INCREASETIME`という`=INC`名前を選択します。</span><span class="sxs-lookup"><span data-stu-id="8fa5d-130">Choose a name like `=INCREASETIME` rather than `=INC`.</span></span>
- <span data-ttu-id="8fa5d-131">類似のアクションを実行する関数に対して同じ動詞を一貫して使用します。</span><span class="sxs-lookup"><span data-stu-id="8fa5d-131">Consistently use the same verbs for functions which perform similar actions.</span></span> <span data-ttu-id="8fa5d-132">たとえば、とで`=DELETEZIPCODE`は`=DELETEADDRESS`なく`=DELETEZIPCODE`を使用し`=REMOVEADDRESS`ます。</span><span class="sxs-lookup"><span data-stu-id="8fa5d-132">For example, use `=DELETEZIPCODE` and `=DELETEADDRESS`, rather than `=DELETEZIPCODE` and `=REMOVEADDRESS`.</span></span>

## <a name="localizing-function-names"></a><span data-ttu-id="8fa5d-133">関数名のローカライズ</span><span class="sxs-lookup"><span data-stu-id="8fa5d-133">Localizing function names</span></span>

<span data-ttu-id="8fa5d-134">個別の JSON ファイルを使用し、アドインのマニフェストファイルで値をオーバーライドすることにより、異なる言語の関数名をローカライズできます。</span><span class="sxs-lookup"><span data-stu-id="8fa5d-134">You can localize your function names for different languages using separate JSON files and override values in your add-in's manifest file.</span></span> <span data-ttu-id="8fa5d-135">これはローカライズされた関数と競合する`id`可能性`name`があるため、ベストプラクティスとして、関数または組み込みの Excel 関数を別の言語で提供しないようにします。</span><span class="sxs-lookup"><span data-stu-id="8fa5d-135">As a best practice, avoid giving your functions an `id` or `name` that is a built-in Excel function in another language as this could conflict with localized functions.</span></span>

<span data-ttu-id="8fa5d-136">ローカライズの詳細については、「[カスタム関数をローカライズ](custom-functions-localize.md)する」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="8fa5d-136">For full information on localizing, see [Localize custom functions](custom-functions-localize.md)</span></span>

## <a name="next-steps"></a><span data-ttu-id="8fa5d-137">次の手順</span><span class="sxs-lookup"><span data-stu-id="8fa5d-137">Next steps</span></span>
<span data-ttu-id="8fa5d-138">[エラー処理のベストプラクティス](custom-functions-errors.md)について説明します。</span><span class="sxs-lookup"><span data-stu-id="8fa5d-138">Learn about [error handling best practices](custom-functions-errors.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="8fa5d-139">関連項目</span><span class="sxs-lookup"><span data-stu-id="8fa5d-139">See also</span></span>

* [<span data-ttu-id="8fa5d-140">カスタム関数のメタデータ</span><span class="sxs-lookup"><span data-stu-id="8fa5d-140">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="8fa5d-141">カスタム関数のベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="8fa5d-141">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="8fa5d-142">チュートリアル: Excel でカスタム関数を作成します。</span><span class="sxs-lookup"><span data-stu-id="8fa5d-142">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
* [<span data-ttu-id="8fa5d-143">Excel カスタム関数のランタイム</span><span class="sxs-lookup"><span data-stu-id="8fa5d-143">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
