---
title: Office アドインのアイコン ガイドライン
description: アドインコマンドのためのアイコンの設計方法と、最新のデザインスタイルおよび Monoline デザインスタイルの概要を説明します。
ms.date: 12/09/2019
localization_priority: Normal
ms.openlocfilehash: b6a960b038b7e02f75101f589469db328465d6bc
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/08/2020
ms.locfileid: "44607674"
---
# <a name="icons"></a><span data-ttu-id="9540a-103">アイコン</span><span class="sxs-lookup"><span data-stu-id="9540a-103">Icons</span></span>

<span data-ttu-id="9540a-104">アイコンは、動作や概念を視覚的に表現するものです。</span><span class="sxs-lookup"><span data-stu-id="9540a-104">Icons are the visual representation of a behavior or concept.</span></span> <span data-ttu-id="9540a-105">多くの場合、コントロールとコマンドに意味を与えるために使用します。</span><span class="sxs-lookup"><span data-stu-id="9540a-105">They are often used to add meaning to controls and commands.</span></span> <span data-ttu-id="9540a-106">環境内でユーザーが移動するのにサインが役立つのと同じように、リアルなビジュアルや象徴的なビジュアルにより、ユーザーは UI 間を移動できるようになります。</span><span class="sxs-lookup"><span data-stu-id="9540a-106">Visuals, either realistic or symbolic, enable the user to navigate the UI the same way signs help users navigate their environment.</span></span> <span data-ttu-id="9540a-107">お客様がコントロールを選択するときの動作をすばやく解析できるにように、必要な詳細のみを含む、シンプルで明確なビジュアルにする必要があります。</span><span class="sxs-lookup"><span data-stu-id="9540a-107">They should be simple, clear, and contain only the necessary details to enable customers to quickly parse what action will occur when they choose a control.</span></span>

<span data-ttu-id="9540a-108">Office リボンのインターフェイスには、標準の視覚的なスタイルが使用されています。</span><span class="sxs-lookup"><span data-stu-id="9540a-108">Office ribbon interfaces have a standard visual style.</span></span> <span data-ttu-id="9540a-109">これにより、Office アプリとの間で一貫性と親和性を保つことができます。</span><span class="sxs-lookup"><span data-stu-id="9540a-109">This ensures consistency and familiarity across Office apps.</span></span> <span data-ttu-id="9540a-110">このガイドラインは、ソリューションの PNG アセットのセットを Office の自然な一部のように設計するのに役立ちます。</span><span class="sxs-lookup"><span data-stu-id="9540a-110">The guidelines will help you design a set of PNG assets for your solution that fit in as a natural part of Office.</span></span>

<span data-ttu-id="9540a-p103">多くの HTML コンテナーには、コントロールとアイコン画像が含まれています。Office UI Fabric のカスタム フォントを使用して、Office スタイルのアイコンをアドインで表示します。Fabric のアイコンのフォントには、ニーズに合わせて拡大縮小、色付け、スタイリングできる一般的な Office メタファーのグリフが多数含まれています。独自のアイコンのセットで既存のビジュアル言語を使用している場合、HTML キャンバスでも自由に使用できます。アイコンの標準セットを使用して独自のブランドの継続性を築くことは、すべてのデザイン言語の重要な一部をなしています。Office メタファーとの競合により、お客様が混乱することのないように注意してください。</span><span class="sxs-lookup"><span data-stu-id="9540a-p103">Many HTML containers contain controls with iconography. Use Office UI Fabric’s custom font to render Office styled icons in your add-in. Fabric’s icon font contains many glyphs for common Office metaphors that you can scale, color, and style to suit your needs. If you have an existing visual language with your own set of icons, feel free to use it in your HTML canvases. Building continuity with your own brand with a standard set of icons is an important part of any design language. Be careful to avoid creating confusion for customers by conflicting with Office metaphors.</span></span>

## <a name="design-icons-for-add-in-commands"></a><span data-ttu-id="9540a-117">アドイン コマンドのアイコンをデザインする</span><span class="sxs-lookup"><span data-stu-id="9540a-117">Design icons for add-in commands</span></span>

<span data-ttu-id="9540a-118">[アドイン コマンド](add-in-commands.md)は、Office UI にボタン、テキスト、およびアイコンを追加します。</span><span class="sxs-lookup"><span data-stu-id="9540a-118">[Add-in commands](add-in-commands.md) add buttons, text, and icons to the Office UI.</span></span> <span data-ttu-id="9540a-119">アドイン コマンドのボタンには、ユーザーがコマンドを使うときに、実行しようとするアクションを明確に識別できる、分かりやすいアイコンとラベルをつける必要があります。</span><span class="sxs-lookup"><span data-stu-id="9540a-119">Your add-in command buttons should provide meaningful icons and labels that clearly identify the action the user is taking when they use a command.</span></span> <span data-ttu-id="9540a-120">次の記事では、Office とシームレスに統合されるアイコンを設計するのに役立つスタイルと運用上のガイドラインを提供します。</span><span class="sxs-lookup"><span data-stu-id="9540a-120">The following articles provide stylistic and production guidelines to help you design icons that integrate seamlessly with Office.</span></span>

- <span data-ttu-id="9540a-121">Monoline スタイルの Office 365 については、「 [Office アドインの Monoline スタイルアイコンのガイドライン](add-in-icons-monoline.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="9540a-121">For the Monoline style of Office 365, see [Monoline style icon guidelines for Office Add-ins](add-in-icons-monoline.md).</span></span>
- <span data-ttu-id="9540a-122">サブスクリプション以外の Office 2013 以降の新しいスタイルについては、「 [Office アドインの新しいスタイルのアイコンガイドライン](add-in-icons-fresh.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="9540a-122">For the Fresh style of non-subscription Office 2013+, see [Fresh style icon guidelines for Office Add-ins](add-in-icons-fresh.md).</span></span>

> [!NOTE]
> <span data-ttu-id="9540a-123">どちらか一方のスタイルを選択する必要があり、アドインは Office 365 またはサブスクリプション以外の Office で実行されている場合と同じアイコンを使用します。</span><span class="sxs-lookup"><span data-stu-id="9540a-123">You must choose one style or the other and your add-in will use the same icons whether it is running in Office 365 or non-subscription Office.</span></span>

## <a name="see-also"></a><span data-ttu-id="9540a-124">関連項目</span><span class="sxs-lookup"><span data-stu-id="9540a-124">See also</span></span>

- [<span data-ttu-id="9540a-125">アドイン開発のベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="9540a-125">Add-in development best practices</span></span>](../concepts/add-in-development-best-practices.md)
- [<span data-ttu-id="9540a-126">Excel、Word、PowerPoint のアドイン コマンド</span><span class="sxs-lookup"><span data-stu-id="9540a-126">Add-in commands for Excel, Word, and PowerPoint</span></span>](../design/add-in-commands.md)
