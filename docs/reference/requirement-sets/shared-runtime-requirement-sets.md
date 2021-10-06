---
title: 共有ランタイム要件セット
description: SharedRuntime API をサポートOfficeするプラットフォームとアプリケーションを指定します。
ms.date: 10/05/2021
ms.prod: non-product-specific
ms.localizationpriority: medium
ms.openlocfilehash: c202ba3047b8dc3998b3bcfd241274bcedf5a93d
ms.sourcegitcommit: 489befc41e543a4fb3c504fd9b3f61322134c1ef
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/06/2021
ms.locfileid: "60138500"
---
# <a name="shared-runtime-requirement-sets"></a>共有ランタイム要件セット

要件セットは、API メンバーの名前付きグループです。Office アドインは、マニフェストで指定されている要件セットを使用するか、ランタイム チェックを使用して、Office アプリケーションがアドインに必要な API をサポートしているかどうかを判別します。詳しくは、「[Office のバージョンと要件セット](../../develop/office-versions-and-requirement-sets.md)」をご覧ください。

作業ウィンドウ、アドイン コマンドから起動される関数ファイル、Excel カスタム関数など、JavaScript コードを実行する Office アドインの一部は、単一の JavaScript ランタイムを共有できます。 これにより、すべてのパーツで一連のグローバル変数を共有したり、読み込まれたライブラリのセットを共有したり、永続ストレージを介してメッセージを渡したりすることなく相互に通信することができます。 詳細については、「共有 JavaScript ランタイムを使用Officeアドインを構成する[」を参照してください](../../develop/configure-your-add-in-to-use-a-shared-runtime.md)。

次の表に、SharedRuntime 1.1 要件セット、その要件セットをサポートする Office クライアント アプリケーション、および Office アプリケーションのビルドまたはバージョン番号を示します。

| 要件セット | Office 2021 以降のWindows<br>(1 回限りの購入) | Windows での Office<br>(Microsoft 365 サブスクリプションに接続) | Office on iPad<br>(Microsoft 365 サブスクリプションに接続) | Office on Mac<br>(Microsoft 365 サブスクリプションに接続) | Office on the web | Office Online Server |
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| SharedRuntime 1.1  | ビルド 16.0.14326.20454 以降 | バージョン 2002 (ビルド 12527.20092) 以降 | 該当なし | 16.35 以降 | 2020 年 2 月 | 該当なし |

> [!IMPORTANT]
> 共有 JavaScript ランタイム要件セットは、次のアプリケーションおよびプラットフォームOfficeでのみ使用できます。
>
> - Excel on the web、Windows、および Mac。
> - Windows 版の PowerPoint (ビルド 13218.10000 以降)。 PowerPoint の共有 JavaScript ランタイムは現在プレビュー中であり、変更される可能性があります。 運用環境での使用はサポートされません。 最新のビルドを入手するには、[Office Insider に参加する](https://insider.office.com/join)必要があります。 プレビュー機能を試す良い方法は、Microsoft 365 サブスクリプションを使用することです。 Microsoft 365 サブスクリプションをまだお持ちでない場合は、[Microsoft 365 開発者プログラム](https://developer.microsoft.com/office/dev-program)に参加することで入手できます。
>
> 現時点では、共有 JavaScript ランタイムは iPad または Office 2019 以前の 1 回限りの購入バージョンではサポートされません。

## <a name="office-versions-and-build-numbers"></a>Office のバージョンとビルド番号

バージョン、ビルド番号、Office Online Server の詳細については以下を参照してください。

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]
- [Office Online Server 概要](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Office 共通 API の要件セット

共通 API の要件セットの詳細については、「[Office 共通 API の要件セット](office-add-in-requirement-sets.md)」をご覧ください。

## <a name="see-also"></a>関連項目

- [Office アドインを構成して共有 JavaScript ランタイムを使用する](../../develop/configure-your-add-in-to-use-a-shared-runtime.md)
- [Office のバージョンと要件セット](../../develop/office-versions-and-requirement-sets.md)
- [Office アプリケーションと API 要件を指定する](../../develop/specify-office-hosts-and-api-requirements.md)
- [Office アドインの XML マニフェスト](../../develop/add-in-manifests.md)
