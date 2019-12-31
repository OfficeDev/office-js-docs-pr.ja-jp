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
# <a name="clear-the-office-cache"></a>Office のキャッシュをクリアする

以前に Windows、Mac、または iOS にサイドロードしたアドインは、コンピューターで Office のキャッシュをクリアすることにより削除できます。 

また、アドインのマニフェストに変更を加えた場合は (アイコンのファイル名やアドイン コマンドのテキストを更新した場合など)、Office のキャッシュをクリアし、更新されたマニフェストを使用してアドインをサイドロードし直す必要があります。 これを実行することにより、アドインは更新されたマニフェストの記載どおりに Office で表示されるようになります。

## <a name="clear-the-office-cache-on-windows"></a>Windows で Office のキャッシュをクリアする

Windows で Office のキャッシュをクリアするには、フォルダー `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\` 内のコンテンツを削除します。

## <a name="clear-the-office-cache-on-mac"></a>Mac で Office のキャッシュをクリアする

[!include[additional cache folders on Mac](../includes/mac-cache-folders.md)]

##  <a name="clear-the-office-cache-on-ios"></a>iOS で Office のキャッシュをクリアする

iOS で Office のキャッシュをクリアするには、アドイン内の JavaScript から `window.location.reload(true)` を呼び出し、強制的に再読み込みを行います。 別の方法として、Office を再インストールすることもできます。

## <a name="see-also"></a>関連項目

- [Office アドインの XML マニフェスト](../develop/add-in-manifests.md)
- [Office アドインのマニフェストを検証する](troubleshoot-manifest.md)
- [ランタイム ログを使用してアドインをデバッグする](runtime-logging.md)
- [テスト用に Office アドインをサイドロードする](sideload-office-add-ins-for-testing.md)
- [Office アドインをデバッグする](debug-add-ins-using-f12-developer-tools-on-windows-10.md)