---
ms.date: 02/08/2019
description: Excel カスタム関数の名前の要件について説明し、一般的な名前付けの落とし穴を回避します。
title: Excel でのカスタム関数の名前付けのガイドライン (プレビュー)
localization_priority: Normal
ms.openlocfilehash: bdf31879fb6e750fb9dea51f66c55dbc83a2dc90
ms.sourcegitcommit: 8e20e7663be2aaa0f7a5436a965324d171bc667d
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/22/2019
ms.locfileid: "30203850"
---
# <a name="naming-guidelines"></a><span data-ttu-id="09321-103">名前付けのガイドライン</span><span class="sxs-lookup"><span data-stu-id="09321-103">Naming guidelines</span></span>

<span data-ttu-id="09321-104">カスタム関数は、JSON メタデータファイルの**id**および**name**プロパティによって識別されます。</span><span class="sxs-lookup"><span data-stu-id="09321-104">A custom function is identified by an **id** and **name** property in the JSON metadata file.</span></span> <span data-ttu-id="09321-105">関数 id は、JavaScript コードのカスタム関数を一意に識別するために使用されます。</span><span class="sxs-lookup"><span data-stu-id="09321-105">The function id is used to uniquely identify custom functions in your JavaScript code.</span></span> <span data-ttu-id="09321-106">関数名は、Excel でユーザーに表示される表示名として使用されます。</span><span class="sxs-lookup"><span data-stu-id="09321-106">The function name is used as the display name that appears to a user in Excel.</span></span> <span data-ttu-id="09321-107">関数の名前は、ローカライズのためなど、関数の ID とは異なる場合があります。</span><span class="sxs-lookup"><span data-stu-id="09321-107">A function name can differ from the function ID, such as for localization purposes.</span></span> <span data-ttu-id="09321-108">しかし、一般的には、それが異なるという説得力のある理由がない場合は、ID と同じままにしておく必要があります。</span><span class="sxs-lookup"><span data-stu-id="09321-108">But in general it should stay the same as the ID if there is no compelling reason for them to differ.</span></span>

<span data-ttu-id="09321-109">関数名と関数 id は、いくつかの一般的な要件を共有します。</span><span class="sxs-lookup"><span data-stu-id="09321-109">Function names and function IDs share some common requirements:</span></span>

- <span data-ttu-id="09321-110">これらの文字は英数字 (Unicode を含む) である必要があります。数字 0 ~ 9、アンダースコア、ピリオドを使用する必要があります。</span><span class="sxs-lookup"><span data-stu-id="09321-110">They must only use alphanumeric characters (including Unicode), the numbers zero through nine, underscores, and periods.</span></span>

- <span data-ttu-id="09321-111">文字で始まる必要があり、最小で3文字に制限されています。</span><span class="sxs-lookup"><span data-stu-id="09321-111">They must start with a letter and have a minimum limit of three characters.</span></span>

<span data-ttu-id="09321-112">Excel は、組み込み関数名 (など`SUM`) に大文字を使用します。</span><span class="sxs-lookup"><span data-stu-id="09321-112">Excel uses uppercase letters for built-in function names (such as `SUM`).</span></span> <span data-ttu-id="09321-113">そのため、ベストプラクティスとして、カスタム関数名と関数 id に大文字を使用することを検討してください。</span><span class="sxs-lookup"><span data-stu-id="09321-113">Therefore, consider using uppercase letters for your custom function names and function IDs as a best practice.</span></span>

<span data-ttu-id="09321-114">関数名には、次のような名前を付けないでください。</span><span class="sxs-lookup"><span data-stu-id="09321-114">Function names shouldn't be named the same as:</span></span>

- <span data-ttu-id="09321-115">A1 から XFD1048576 のセル、または R1C1 から R1048576C16384 までのセル。</span><span class="sxs-lookup"><span data-stu-id="09321-115">Any cells between A1 to XFD1048576 or any cells between R1C1 to R1048576C16384.</span></span>

- <span data-ttu-id="09321-116">任意の Excel 4.0 マクロ関数 ( `RUN`、 `ECHO`など)。</span><span class="sxs-lookup"><span data-stu-id="09321-116">Any Excel 4.0 Macro Function (such as `RUN`, `ECHO`).</span></span>  <span data-ttu-id="09321-117">これらの関数の完全な一覧については、[この記事](https://www.microsoft.com/en-us/download/details.aspx?id=1465)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="09321-117">For a full list of these functions, see [this article](https://www.microsoft.com/en-us/download/details.aspx?id=1465).</span></span>

## <a name="naming-conflicts"></a><span data-ttu-id="09321-118">名前付けの競合</span><span class="sxs-lookup"><span data-stu-id="09321-118">Naming conflicts</span></span>

<span data-ttu-id="09321-119">関数名が既に存在するアドインの関数名と同じである場合、 **#REF!**</span><span class="sxs-lookup"><span data-stu-id="09321-119">If your function name is the same as a function name in an add-in that already exists, the **#REF!**</span></span> <span data-ttu-id="09321-120">エラーがブックに表示されます。</span><span class="sxs-lookup"><span data-stu-id="09321-120">error will appear in your workbook.</span></span>

<span data-ttu-id="09321-121">名前の競合を修正するには、アドイン内の名前を変更して、関数を再度実行します。</span><span class="sxs-lookup"><span data-stu-id="09321-121">To fix a name conflict, change the name in your add-in and try the function again.</span></span> <span data-ttu-id="09321-122">競合する名前を使用してアドインをアンインストールすることもできます。</span><span class="sxs-lookup"><span data-stu-id="09321-122">You can also uninstall the add-in with the conflicting name.</span></span> <span data-ttu-id="09321-123">または、別の環境でアドインをテストしている場合は、別の名前空間を使用して、関数を区別します (NAMESPACE_NAMEOFFUNCTION など)。</span><span class="sxs-lookup"><span data-stu-id="09321-123">Or, if you're testing your add-in in different environments, try using a different namespace to differentiate your function (such as NAMESPACE_NAMEOFFUNCTION).</span></span>

<span data-ttu-id="09321-124">また、アドイン内で関数を使用する方法についても検討します。</span><span class="sxs-lookup"><span data-stu-id="09321-124">Also consider how you'd like people to use the functions within your add-in.</span></span> <span data-ttu-id="09321-125">多くの場合、同じまたは似た名前を持つ複数の関数を作成するのではなく、複数の引数を関数に追加することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="09321-125">In many cases, it makes sense to add multiple arguments to a function rather than create multiple functions with the same or similar names.</span></span>

## <a name="see-also"></a><span data-ttu-id="09321-126">関連項目</span><span class="sxs-lookup"><span data-stu-id="09321-126">See also</span></span>

* [<span data-ttu-id="09321-127">カスタム関数のメタデータ</span><span class="sxs-lookup"><span data-stu-id="09321-127">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="09321-128">カスタム関数のベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="09321-128">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="09321-129">チュートリアル: Excel でカスタム関数を作成します。</span><span class="sxs-lookup"><span data-stu-id="09321-129">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
* [<span data-ttu-id="09321-130">Excel カスタム関数のランタイム</span><span class="sxs-lookup"><span data-stu-id="09321-130">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
