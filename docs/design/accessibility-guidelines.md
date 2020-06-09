---
title: Office アドインのアクセシビリティ ガイドライン
description: すべてのユーザーが Office アドインにアクセスできるようにする方法について説明します。
ms.date: 09/24/2018
localization_priority: Normal
ms.openlocfilehash: 889563af8ab5f7bbcd4037eedb42933369a92cf2
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/08/2020
ms.locfileid: "44607993"
---
# <a name="accessibility-guidelines"></a><span data-ttu-id="a3bb7-103">アクセシビリティ ガイドライン</span><span class="sxs-lookup"><span data-stu-id="a3bb7-103">Accessibility guidelines</span></span>

<span data-ttu-id="a3bb7-p101">Office アドインを設計して開発する際は、アドインを使用する可能性のあるすべてのユーザーおよび顧客が正常に使用できるものにするよう努める必要があります。ソリューションを、すべての対象ユーザーがアクセス可能なものにするためには、次のガイドラインを適用します。</span><span class="sxs-lookup"><span data-stu-id="a3bb7-p101">As you design and develop your Office Add-ins, you'll want to ensure that all potential users and customers are able to use your add-in successfully. Apply the following guidelines to ensure that your solution is accessible to all audiences.</span></span>

## <a name="design-for-multiple-input-methods"></a><span data-ttu-id="a3bb7-106">複数の入力方法の設計</span><span class="sxs-lookup"><span data-stu-id="a3bb7-106">Design for multiple input methods</span></span>

- <span data-ttu-id="a3bb7-p102">ユーザーがキーボードのみを使用して操作を実行できることを確認します。ユーザーは、Tab キーと矢印キーの組み合わせを使用して、ページ上のすべての実行可能な要素に移動できる必要があります。</span><span class="sxs-lookup"><span data-stu-id="a3bb7-p102">Ensure that users can perform operations by using only the keyboard. Users should be able to move to all actionable elements on the page by using a combination of the Tab and arrow keys.</span></span>
- <span data-ttu-id="a3bb7-109">モバイル デバイスでは、ユーザーがタッチでコントロールを操作するとき、デバイスが便利なオーディオ フィードバックを出すようにします。</span><span class="sxs-lookup"><span data-stu-id="a3bb7-109">On a mobile device, when users operate a control by touch, the device should provide useful audio feedback.</span></span>
- <span data-ttu-id="a3bb7-110">すべての対話型コントロールに、わかりやすいラベルを付けます。</span><span class="sxs-lookup"><span data-stu-id="a3bb7-110">Provide helpful labels for all interactive controls.</span></span> 

## <a name="make-your-add-in-easy-to-use"></a><span data-ttu-id="a3bb7-111">アドインを使いやすいようにする</span><span class="sxs-lookup"><span data-stu-id="a3bb7-111">Make your add-in easy to use</span></span>

- <span data-ttu-id="a3bb7-112">UI 内での意味を伝えるために、色、サイズ、図形、位置、向き、またはサウンドなどの 1 つの属性だけに依存しないようにします。</span><span class="sxs-lookup"><span data-stu-id="a3bb7-112">Don't rely on a single attribute, such as color, size, shape, location, orientation, or sound, to convey meaning in your UI.</span></span>
- <span data-ttu-id="a3bb7-113">ユーザーが操作しないのに別の UI 要素にフォーカスを移動するなどの、コンテキストの予期しない変更を避けます。</span><span class="sxs-lookup"><span data-stu-id="a3bb7-113">Avoid unexpected changes of context, such as moving the focus to a different UI element without user action.</span></span>
- <span data-ttu-id="a3bb7-114">すべてのバインディング操作について、それを検証、確認、取り消す方法を提供します。</span><span class="sxs-lookup"><span data-stu-id="a3bb7-114">Provide a way to verify, confirm, or reverse all binding actions.</span></span>
- <span data-ttu-id="a3bb7-115">オーディオやビデオなどのメディアを一時停止または停止する方法を提供します。</span><span class="sxs-lookup"><span data-stu-id="a3bb7-115">Provide a way to pause or stop media, such as audio and video.</span></span>
- <span data-ttu-id="a3bb7-116">ユーザー操作の時間制限を設けないようにします。</span><span class="sxs-lookup"><span data-stu-id="a3bb7-116">Do not impose a time limit for user action.</span></span>

## <a name="make-your-add-in-easy-to-see"></a><span data-ttu-id="a3bb7-117">アドインを見やすいものにする</span><span class="sxs-lookup"><span data-stu-id="a3bb7-117">Make your add-in easy to see</span></span>

- <span data-ttu-id="a3bb7-118">予期しない色の変更は避けます。</span><span class="sxs-lookup"><span data-stu-id="a3bb7-118">Avoid unexpected color changes.</span></span>
- <span data-ttu-id="a3bb7-p103">UI 要素、タイトルと見出し、入力、エラーを説明する情報を、意味のあるタイムリーなしかたで提供します。コントロールの名前は、そのコントロールの目的を適切に表すものにします。</span><span class="sxs-lookup"><span data-stu-id="a3bb7-p103">Provide meaningful and timely information to describe UI elements, titles and headings, inputs, and errors. Ensure that names of controls adequately describe the intent of the control.</span></span>
- <span data-ttu-id="a3bb7-121">色のコントラストについては、[標準ガイドライン](https://www.w3.org/TR/UNDERSTANDING-WCAG20/visual-audio-contrast-contrast.html)に従います。</span><span class="sxs-lookup"><span data-stu-id="a3bb7-121">Follow [standard guidelines](https://www.w3.org/TR/UNDERSTANDING-WCAG20/visual-audio-contrast-contrast.html) for color contrast.</span></span>

## <a name="account-for-assistive-technologies"></a><span data-ttu-id="a3bb7-122">支援テクノロジ用のアカウント</span><span class="sxs-lookup"><span data-stu-id="a3bb7-122">Account for assistive technologies</span></span>

- <span data-ttu-id="a3bb7-123">ビジュアル、オーディオ、その他の対話式操作を含め、支援テクノロジの妨げになる機能を使用しないようにします。</span><span class="sxs-lookup"><span data-stu-id="a3bb7-123">Avoid using features that interfere with assistive technologies, including visual, audio, or other interactions.</span></span>
- <span data-ttu-id="a3bb7-p104">テキストをイメージ形式で提供しないようにします。スクリーン リーダーは、イメージ内のテキストを読み取ることができません。</span><span class="sxs-lookup"><span data-stu-id="a3bb7-p104">Do not provide text in an image format. Screen readers cannot read text within images.</span></span>
- <span data-ttu-id="a3bb7-126">すべてのオーディオ ソースを調整またはミュートする方法をユーザーに提供します。</span><span class="sxs-lookup"><span data-stu-id="a3bb7-126">Provide a way for users to adjust or mute all audio sources.</span></span>
- <span data-ttu-id="a3bb7-127">キャプションまたはオーディオ ソースによるオーディオ説明を有効にする方法をユーザーに提供します。</span><span class="sxs-lookup"><span data-stu-id="a3bb7-127">Provide a way for users to turn on captions or audio description with audio sources.</span></span>
- <span data-ttu-id="a3bb7-128">ユーザーに警告する手段として、視覚的な合図や振動など、サウンドに代わる方法を提供します。</span><span class="sxs-lookup"><span data-stu-id="a3bb7-128">Provide alternatives to sound as a means to alert users, such as visual cues or vibrations.</span></span>

## <a name="see-also"></a><span data-ttu-id="a3bb7-129">関連項目</span><span class="sxs-lookup"><span data-stu-id="a3bb7-129">See also</span></span>

- [<span data-ttu-id="a3bb7-130">Web コンテンツ アクセシビリティ ガイドライン (WCAG) 2.0</span><span class="sxs-lookup"><span data-stu-id="a3bb7-130">Web Content Accessibility Guidelines (WCAG) 2.0</span></span>](https://www.w3.org/TR/wcag2ict/#REF-WCAG20)
- [<span data-ttu-id="a3bb7-131">WCAG 2.0 の非 Web 情報および通信テクノロジへの適用ガイダンス (WCAG2ICT)</span><span class="sxs-lookup"><span data-stu-id="a3bb7-131">Guidance on Applying WCAG 2.0 to Non-Web Information and Communications Technologies (WCAG2ICT)</span></span>](https://www.w3.org/TR/wcag2ict/)
- [<span data-ttu-id="a3bb7-132">情報および通信テクノロジ (ICT) におけるユーザー補助要件の欧州基準</span><span class="sxs-lookup"><span data-stu-id="a3bb7-132">European Standard on accessibility requirements for Information and Communication Technologies (ICT)</span></span>](https://www.etsi.org/deliver/etsi_en/301500_301599/301549/01.00.00_20/en_301549v010000c.pdf) 
