---
title: Office アドインのデザイン言語
description: アドインと視覚的に互換性Officeする方法について説明Office。
ms.date: 05/12/2021
localization_priority: Normal
ms.openlocfilehash: 4374eaad1e660728a347d19a012d345b642755af
ms.sourcegitcommit: 30f6c620380075e3459cac748ca0c656427b384d
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/12/2021
ms.locfileid: "52330060"
---
# <a name="office-add-in-design-language"></a><span data-ttu-id="584ac-103">Office アドインのデザイン言語</span><span class="sxs-lookup"><span data-stu-id="584ac-103">Office Add-in design language</span></span>

<span data-ttu-id="584ac-p101">Office のデザイン言語は、エクスペリエンス全体の整合性を保証するクリーンでシンプルなビジュアル システムです。Office のインターフェイスを定義する、次のようなビジュアル要素のセットが含まれています。</span><span class="sxs-lookup"><span data-stu-id="584ac-p101">The Office design language is a clean and simple visual system that ensures consistency across experiences. It contains a set of visual elements that define Office interfaces, including:</span></span>

- <span data-ttu-id="584ac-106">標準的な書体</span><span class="sxs-lookup"><span data-stu-id="584ac-106">A standard typeface</span></span>
- <span data-ttu-id="584ac-107">一般的なカラー パレット</span><span class="sxs-lookup"><span data-stu-id="584ac-107">A common color palette</span></span>
- <span data-ttu-id="584ac-108">文字体裁のサイズと太さのセット</span><span class="sxs-lookup"><span data-stu-id="584ac-108">A set of typographic sizes and weights</span></span>
- <span data-ttu-id="584ac-109">アイコンのガイドライン</span><span class="sxs-lookup"><span data-stu-id="584ac-109">Icon guidelines</span></span>
- <span data-ttu-id="584ac-110">共有アイコンのアセット</span><span class="sxs-lookup"><span data-stu-id="584ac-110">Shared icon assets</span></span>
- <span data-ttu-id="584ac-111">アニメーションの定義</span><span class="sxs-lookup"><span data-stu-id="584ac-111">Animation definitions</span></span>
- <span data-ttu-id="584ac-112">一般的なコンポーネント</span><span class="sxs-lookup"><span data-stu-id="584ac-112">Common components</span></span>

<span data-ttu-id="584ac-113">[Fluent UI](../design/add-in-design.md)は、デザイン言語を使用して構築するためのOfficeフレームワークです。</span><span class="sxs-lookup"><span data-stu-id="584ac-113">[Fluent UI](../design/add-in-design.md) is the official front-end framework for building with the Office design language.</span></span> <span data-ttu-id="584ac-114">Fluent UI の使用はオプションですが、アドインが自然な拡張機能のように感じOffice。</span><span class="sxs-lookup"><span data-stu-id="584ac-114">Using Fluent UI is optional, but it is the fastest way to ensure that your add-ins feel like a natural extension of Office.</span></span> <span data-ttu-id="584ac-115">Fluent UI を活用して、ユーザーを補完するアドインを設計およびOffice。</span><span class="sxs-lookup"><span data-stu-id="584ac-115">Take advantage of Fluent UI to design and build add-ins that complement Office.</span></span>

<span data-ttu-id="584ac-p103">多くの Office アドインは、以前から存在するブランドに関連付けられています。強力なブランドとそのビジュアルまたはコンポーネント言語を、アドインに保持できます。Office と統合する際に、独自のビジュアル言語を保持する機会を探します。Office の色、文字体裁、アイコン、その他のスタイルの要素を、独自のブランドの要素と交換する方法を検討してください。お客様によく知られているコントロールやコンポーネントを挿入する際は、一般的なアドイン レイアウトや UX 設計パターンに従う方法を検討してください。</span><span class="sxs-lookup"><span data-stu-id="584ac-p103">Many Office Add-ins are associated with a preexisting brand. You can retain a strong brand and its visual or component language in your add-in. Look for opportunities to retain your own visual language while integrating with Office. Consider ways to swap out Office colors, typography, icons, or other stylistic elements with elements of your own brand. Consider ways to follow common add-in layouts or UX design patterns while inserting controls and components that are familiar to your customers.</span></span>

<span data-ttu-id="584ac-p104">過度にブランド化された HTML ベースの UI を Office 内に挿入すると、お客様が不満を抱く可能性があります。Office にシームレスに適合するためのバランスを見つけるだけでなく、サービスや親ブランドとはっきり調和するようにします。アドインが Office に適合しないとき、ほとんどの場合はスタイル要素の競合が原因です。たとえば、文字体裁が大きすぎてグリッド線を越えている、色が対照的で派手である、アニメーションが余計で Office と動作が異なる場合です。コントロールやコンポーネントの外観と動作が、Office の基準から大幅にそれています。</span><span class="sxs-lookup"><span data-stu-id="584ac-p104">Inserting a heavily branded HTML-based UI inside of Office can create dissonance for customers. Find a balance that fits seamlessly in Office but also clearly aligns with your service or parent brand. When an add-in does not fit with Office, it's often because stylistic elements conflict. For example, typography is too large and off grid, colors are contrasting or particularly loud, or animations are superfluous and behave differently than Office. The appearance and behavior of controls or components veer too far from Office standards.</span></span>
