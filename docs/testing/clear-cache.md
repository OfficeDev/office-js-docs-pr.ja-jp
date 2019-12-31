---
title: Office のキャッシュをクリアする
description: コンピューターで Office のキャッシュをクリアする方法について説明します。
ms.date: 12/31/2019
localization_priority: Priority
ms.openlocfilehash: 3744d8125a5165569c262dc28622614853798c6f
ms.sourcegitcommit: d5ac9284d1e96dc91a9168d7641e44d88535e1a7
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/31/2019
ms.locfileid: "40915072"
---
# <a name="clear-the-office-cache"></a><span data-ttu-id="5cbcd-103">Office のキャッシュをクリアする</span><span class="sxs-lookup"><span data-stu-id="5cbcd-103">Clear the Office cache</span></span>

<span data-ttu-id="5cbcd-104">以前に Windows、Mac、または iOS にサイドロードしたアドインは、コンピューターで Office のキャッシュをクリアすることにより削除できます。</span><span class="sxs-lookup"><span data-stu-id="5cbcd-104">You can remove an add-in that you've previously sideloaded on Windows, Mac, or iOS by clearing the Office cache on your computer.</span></span> 

<span data-ttu-id="5cbcd-105">また、アドインのマニフェストに変更を加えた場合は (アイコンのファイル名やアドイン コマンドのテキストを更新した場合など)、Office のキャッシュをクリアし、更新されたマニフェストを使用してアドインをサイドロードし直す必要があります。</span><span class="sxs-lookup"><span data-stu-id="5cbcd-105">Additionally, if you make changes to your add-in's manifest (for example, update file names of icons or text of add-in commands), you should clear the Office cache and then re-sideload the add-in using updated manifest.</span></span> <span data-ttu-id="5cbcd-106">これを実行することにより、アドインは更新されたマニフェストの記載どおりに Office で表示されるようになります。</span><span class="sxs-lookup"><span data-stu-id="5cbcd-106">Doing so will allow Office to render the add-in as it's described by the updated manifest.</span></span>

## <a name="clear-the-office-cache-on-windows"></a><span data-ttu-id="5cbcd-107">Windows で Office のキャッシュをクリアする</span><span class="sxs-lookup"><span data-stu-id="5cbcd-107">Clear the Office cache on Windows</span></span>

<span data-ttu-id="5cbcd-108">Windows で Office のキャッシュをクリアするには、フォルダー `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\` 内のコンテンツを削除します。</span><span class="sxs-lookup"><span data-stu-id="5cbcd-108">To clear the Office cache on Windows, delete the contents of the folder `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`.</span></span>

## <a name="clear-the-office-cache-on-mac"></a><span data-ttu-id="5cbcd-109">Mac で Office のキャッシュをクリアする</span><span class="sxs-lookup"><span data-stu-id="5cbcd-109">Clear the Office cache on Mac</span></span>

[!include[additional cache folders on Mac](../includes/mac-cache-folders.md)]

##  <a name="clear-the-office-cache-on-ios"></a><span data-ttu-id="5cbcd-110">iOS で Office のキャッシュをクリアする</span><span class="sxs-lookup"><span data-stu-id="5cbcd-110">Clear the Office cache on iOS</span></span>

<span data-ttu-id="5cbcd-111">iOS で Office のキャッシュをクリアするには、アドイン内の JavaScript から `window.location.reload(true)` を呼び出し、強制的に再読み込みを行います。</span><span class="sxs-lookup"><span data-stu-id="5cbcd-111">To clear the Office cache on iOS, call `window.location.reload(true)` from JavaScript in the add-in to force a reload.</span></span> <span data-ttu-id="5cbcd-112">別の方法として、Office を再インストールすることもできます。</span><span class="sxs-lookup"><span data-stu-id="5cbcd-112">Alternatively, you can reinstall Office.</span></span>

## <a name="see-also"></a><span data-ttu-id="5cbcd-113">関連項目</span><span class="sxs-lookup"><span data-stu-id="5cbcd-113">See also</span></span>

- [<span data-ttu-id="5cbcd-114">Office アドインの XML マニフェスト</span><span class="sxs-lookup"><span data-stu-id="5cbcd-114">Office Add-ins XML manifest</span></span>](../develop/add-in-manifests.md)
- [<span data-ttu-id="5cbcd-115">Office アドインのマニフェストを検証する</span><span class="sxs-lookup"><span data-stu-id="5cbcd-115">Validate an Office Add-in manifest</span></span>](troubleshoot-manifest.md)
- [<span data-ttu-id="5cbcd-116">ランタイム ログを使用してアドインをデバッグする</span><span class="sxs-lookup"><span data-stu-id="5cbcd-116">Debug your add-in with runtime logging</span></span>](runtime-logging.md)
- [<span data-ttu-id="5cbcd-117">テスト用に Office アドインをサイドロードする</span><span class="sxs-lookup"><span data-stu-id="5cbcd-117">Sideload Office Add-ins for testing</span></span>](sideload-office-add-ins-for-testing.md)
- [<span data-ttu-id="5cbcd-118">Office アドインをデバッグする</span><span class="sxs-lookup"><span data-stu-id="5cbcd-118">Debug Office Add-ins</span></span>](debug-add-ins-using-f12-developer-tools-on-windows-10.md)