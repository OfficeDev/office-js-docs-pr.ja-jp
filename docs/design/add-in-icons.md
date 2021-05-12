---
title: Office アドインのアイコン ガイドライン
description: アドイン コマンドのアイコンと Fresh および Monoline デザイン スタイルを設計する方法の概要を説明します。
ms.date: 05/12/2021
localization_priority: Normal
ms.openlocfilehash: 3073472332a31688676fba796dccd9920a49581d
ms.sourcegitcommit: 30f6c620380075e3459cac748ca0c656427b384d
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/12/2021
ms.locfileid: "52329969"
---
# <a name="icons"></a><span data-ttu-id="09cf2-103">アイコン</span><span class="sxs-lookup"><span data-stu-id="09cf2-103">Icons</span></span>

<span data-ttu-id="09cf2-104">アイコンは、動作や概念を視覚的に表現するものです。</span><span class="sxs-lookup"><span data-stu-id="09cf2-104">Icons are the visual representation of a behavior or concept.</span></span> <span data-ttu-id="09cf2-105">多くの場合、コントロールとコマンドに意味を与えるために使用します。</span><span class="sxs-lookup"><span data-stu-id="09cf2-105">They are often used to add meaning to controls and commands.</span></span> <span data-ttu-id="09cf2-106">環境内でユーザーが移動するのにサインが役立つのと同じように、リアルなビジュアルや象徴的なビジュアルにより、ユーザーは UI 間を移動できるようになります。</span><span class="sxs-lookup"><span data-stu-id="09cf2-106">Visuals, either realistic or symbolic, enable the user to navigate the UI the same way signs help users navigate their environment.</span></span> <span data-ttu-id="09cf2-107">お客様がコントロールを選択するときの動作をすばやく解析できるにように、必要な詳細のみを含む、シンプルで明確なビジュアルにする必要があります。</span><span class="sxs-lookup"><span data-stu-id="09cf2-107">They should be simple, clear, and contain only the necessary details to enable customers to quickly parse what action will occur when they choose a control.</span></span>

<span data-ttu-id="09cf2-108">Office アプリインターフェイスには標準の表示スタイルがあります。</span><span class="sxs-lookup"><span data-stu-id="09cf2-108">Office app ribbon interfaces have a standard visual style.</span></span> <span data-ttu-id="09cf2-109">これにより、Office アプリとの間で一貫性と親和性を保つことができます。</span><span class="sxs-lookup"><span data-stu-id="09cf2-109">This ensures consistency and familiarity across Office apps.</span></span> <span data-ttu-id="09cf2-110">このガイドラインは、ソリューションの PNG アセットのセットを Office の自然な一部のように設計するのに役立ちます。</span><span class="sxs-lookup"><span data-stu-id="09cf2-110">The guidelines will help you design a set of PNG assets for your solution that fit in as a natural part of Office.</span></span>

<span data-ttu-id="09cf2-111">多くの HTML コンテナーには、コントロールとアイコン画像が含まれています。</span><span class="sxs-lookup"><span data-stu-id="09cf2-111">Many HTML containers contain controls with iconography.</span></span> <span data-ttu-id="09cf2-112">Fabric Core のカスタム フォントを使用して、Officeスタイルのアイコンをレンダリングします。</span><span class="sxs-lookup"><span data-stu-id="09cf2-112">Use Fabric Core’s custom font to render Office styled icons in your add-in.</span></span> <span data-ttu-id="09cf2-113">[Fabric Core](fabric-core.md)によって提供されるアイコン フォントには、ニーズに合わせてスケール、色、およびスタイルを調整Office一般的なメタファー用の多くのグリフが含まれている。</span><span class="sxs-lookup"><span data-stu-id="09cf2-113">The icon font provided by [Fabric Core](fabric-core.md) contains many glyphs for common Office metaphors that you can scale, color, and style to suit your needs.</span></span> <span data-ttu-id="09cf2-114">独自のアイコンのセットで既存のビジュアル言語を使用している場合、HTML キャンバスでも自由に使用できます。</span><span class="sxs-lookup"><span data-stu-id="09cf2-114">If you have an existing visual language with your own set of icons, feel free to use it in your HTML canvases.</span></span> <span data-ttu-id="09cf2-115">アイコンの標準セットを使用して独自のブランドの継続性を築くことは、すべてのデザイン言語の重要な一部をなしています。</span><span class="sxs-lookup"><span data-stu-id="09cf2-115">Building continuity with your own brand with a standard set of icons is an important part of any design language.</span></span> <span data-ttu-id="09cf2-116">Office メタファーとの競合により、お客様が混乱することのないように注意してください。</span><span class="sxs-lookup"><span data-stu-id="09cf2-116">Be careful to avoid creating confusion for customers by conflicting with Office metaphors.</span></span>

## <a name="design-icons-for-add-in-commands"></a><span data-ttu-id="09cf2-117">アドイン コマンドのアイコンをデザインする</span><span class="sxs-lookup"><span data-stu-id="09cf2-117">Design icons for add-in commands</span></span>

<span data-ttu-id="09cf2-118">[アドイン コマンド](add-in-commands.md)は、Office UI にボタン、テキスト、およびアイコンを追加します。</span><span class="sxs-lookup"><span data-stu-id="09cf2-118">[Add-in commands](add-in-commands.md) add buttons, text, and icons to the Office UI.</span></span> <span data-ttu-id="09cf2-119">アドイン コマンドのボタンには、ユーザーがコマンドを使うときに、実行しようとするアクションを明確に識別できる、分かりやすいアイコンとラベルをつける必要があります。</span><span class="sxs-lookup"><span data-stu-id="09cf2-119">Your add-in command buttons should provide meaningful icons and labels that clearly identify the action the user is taking when they use a command.</span></span> <span data-ttu-id="09cf2-120">次の記事では、アプリとシームレスに統合するアイコンを設計するのに役立つ、定型的なガイドラインと運用ガイドラインをOffice。</span><span class="sxs-lookup"><span data-stu-id="09cf2-120">The following articles provide stylistic and production guidelines to help you design icons that integrate seamlessly with Office.</span></span>

- <span data-ttu-id="09cf2-121">[モノライン] のスタイルについてはMicrosoft 365アドインの[モノOfficeガイドラインを参照してください](add-in-icons-monoline.md)。</span><span class="sxs-lookup"><span data-stu-id="09cf2-121">For the Monoline style of Microsoft 365, see [Monoline style icon guidelines for Office Add-ins](add-in-icons-monoline.md).</span></span>
- <span data-ttu-id="09cf2-122">2013 以上のサブスクリプション以外の新しいOfficeについては、「新しいスタイル のアイコン ガイドライン」を参照Office[してください](add-in-icons-fresh.md)。</span><span class="sxs-lookup"><span data-stu-id="09cf2-122">For the Fresh style of non-subscription Office 2013+, see [Fresh style icon guidelines for Office Add-ins](add-in-icons-fresh.md).</span></span>

> [!NOTE]
> <span data-ttu-id="09cf2-123">1 つのスタイルを選択するか、別のスタイルを選択する必要があります。アドインは、Microsoft 365 またはサブスクリプション以外のアプリケーションで実行Office。</span><span class="sxs-lookup"><span data-stu-id="09cf2-123">You must choose one style or the other and your add-in will use the same icons whether it is running in Microsoft 365 or non-subscription Office.</span></span>

## <a name="see-also"></a><span data-ttu-id="09cf2-124">関連項目</span><span class="sxs-lookup"><span data-stu-id="09cf2-124">See also</span></span>

- [<span data-ttu-id="09cf2-125">アドイン開発のベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="09cf2-125">Add-in development best practices</span></span>](../concepts/add-in-development-best-practices.md)
- [<span data-ttu-id="09cf2-126">Excel、Word、PowerPoint のアドイン コマンド</span><span class="sxs-lookup"><span data-stu-id="09cf2-126">Add-in commands for Excel, Word, and PowerPoint</span></span>](../design/add-in-commands.md)
