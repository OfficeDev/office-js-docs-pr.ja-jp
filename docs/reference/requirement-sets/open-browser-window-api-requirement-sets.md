---
title: ブラウザーウィンドウの要件セットを開く
description: OpenBrowserWindow API をサポートする Office プラットフォームとビルドを指定します。
ms.date: 09/16/2020
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: 8bc26525bf64ed87d46d85cd1248f79696d67f2b
ms.sourcegitcommit: 4a03d8b3f676ee2d91114813cb81bce5da3c8d6b
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/22/2020
ms.locfileid: "48175508"
---
# <a name="open-browser-window-api-requirement-sets"></a>ブラウザーウィンドウ API の要件セットを開く

要件セットは、API メンバーの名前付きグループです。Office アドインは、マニフェストで指定されている要件セットを使用するか、ランタイム チェックを使用して、Office ホストがアドインに必要な API をサポートしているかどうかを判別します。詳しくは、「[Office のバージョンと要件セット](../../develop/office-versions-and-requirement-sets.md)」をご覧ください。

OpenBrowserWindow API セットを使用すると、アドイン自体でサンドボックス内の webview コントロールでは実行できないタスクを実行するために、ブラウザーを開いておくことができます。たとえば、Microsoft Edge で webview コントロールが提供されている場合は、PDF ファイルをダウンロードします。

Office アドインは Office の複数のバージョンで機能します。 次の表に、OpenBrowserWindow API の要件セット、その要件セットをサポートする Office ホストアプリケーション、Office アプリケーションのビルド番号またはバージョン番号を示します。

|  要件セット  | Windows 以降の Office 2013<br>(1 回限りの購入) | Windows での Office<br>(Office 365 サブスクリプションに接続) |  Office on iPad<br>(Office 365 サブスクリプションに接続)  |  Office on Mac<br>(Office 365 サブスクリプションに接続)  | Office on the web  |  Office Online Server  |
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|
| OpenBrowserWindowApi 1.1  | N/A | バージョン 1810 (ビルド 16.0.11001.20074) 以降 | 16.0.0.0 以降 | 16.0.0.0 以降 | N/A | N/A|

バージョン、ビルド番号、Office Online Server の詳細については以下を参照してください。

- [Office 365 クライアントの更新プログラム チャネル リリースのバージョン番号およびビルド番号](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [使用している Office のバージョンを確認する方法](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)
- [Office 365 クライアント アプリケーションのバージョン番号およびビルド番号を確認することができます。](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [Office Online Server 概要](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Office 共通 API の要件セット

共通 API の要件セットの詳細については、「[Office 共通 API の要件セット](office-add-in-requirement-sets.md)」をご覧ください。

## <a name="openbrowserwindowapi-11"></a>OpenBrowserWindowApi 1.1

OpenBrowserWindowApi 1.1 は、API の最初のバージョンです。 API の詳細については、「 [Office. ui](/javascript/api/office/office.context#ui) リファレンス」のトピックを参照してください。

## <a name="see-also"></a>関連項目

- [Office のバージョンと要件セット](../../develop/office-versions-and-requirement-sets.md)
- [Office のホストと API の要件を指定する](../../develop/specify-office-hosts-and-api-requirements.md)
- [Office アドインの XML マニフェスト](../../develop/add-in-manifests.md)
