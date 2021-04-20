---
title: ブラウザー ウィンドウの要件セットを開く
description: openBrowserWindow API Officeサポートするプラットフォームとビルドを指定します。
ms.date: 02/09/2021
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: 759c8265b27fab4589e68fe3f2f90a2a283ef005
ms.sourcegitcommit: ccc0a86d099ab4f5ef3d482e4ae447c3f9b818a3
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/14/2021
ms.locfileid: "50237960"
---
# <a name="open-browser-window-api-requirement-sets"></a>ブラウザー ウィンドウ API の要件セットを開く

要件セットは、API メンバーの名前付きグループです。Office アドインは、マニフェストで指定されている要件セットを使用するか、ランタイム チェックを使用して、Office ホストがアドインに必要な API をサポートしているかどうかを判別します。詳しくは、「[Office のバージョンと要件セット](../../develop/office-versions-and-requirement-sets.md)」をご覧ください。

OpenBrowserWindow API セットを使用すると、アドインはブラウザーを開き、アドイン自体のサンドボックス Web ビュー コントロールで必ずしも実行できないタスクを実行できます。たとえば、Webview コントロールが Microsoft Edge によって提供されている場合に PDF ファイルをダウンロードします。

Office アドインは Office の複数のバージョンで機能します。 次の表に、OpenBrowserWindow API の要件セット、その要件セットをサポートする Office ホスト アプリケーション、および Office アプリケーションのビルド番号またはバージョン番号を示します。

|  要件セット  | Office 2013 Windows 以降の場合<br>(1 回限りの購入) | Windows での Office<br>(Microsoft 365 サブスクリプションに接続) |  Office on iPad<br>(Microsoft 365 サブスクリプションに接続)  |  Office on Mac<br>(Microsoft 365 サブスクリプションに接続)  | Office on the web  |  Office Online Server  |
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|
| OpenBrowserWindowApi 1.1  | 該当せず | バージョン 1810 (ビルド 16.0.11001.20074) 以降 | 16.0.0.0 以降 | 16.0.0.0 以降 | 該当なし | 該当なし|

バージョン、ビルド番号、Office Online Server の詳細については以下を参照してください。

- [Microsoft 365 Apps の更新プログラム チャネル リリースのバージョン番号とビルド番号](/officeupdates/update-history-microsoft365-apps-by-date)
- [使用している Office のバージョンを確認する方法](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)
- [クライアント アプリケーションのバージョンとビルド番号Office確認できます。](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [Office Online Server 概要](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Office 共通 API の要件セット

共通 API の要件セットの詳細については、「[Office 共通 API の要件セット](office-add-in-requirement-sets.md)」をご覧ください。

## <a name="openbrowserwindowapi-11"></a>OpenBrowserWindowApi 1.1

OpenBrowserWindowApi 1.1 は、API の最初のバージョンです。 API の詳細については [、Office.context.ui リファレンス トピックを](/javascript/api/office/office.context#ui) 参照してください。

## <a name="see-also"></a>関連項目

- [Office のバージョンと要件セット](../../develop/office-versions-and-requirement-sets.md)
- [Office のホストと API の要件を指定する](../../develop/specify-office-hosts-and-api-requirements.md)
- [Office アドインの XML マニフェスト](../../develop/add-in-manifests.md)
