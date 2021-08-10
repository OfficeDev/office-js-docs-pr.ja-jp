---
title: マニフェスト ファイルの Hosts 要素
description: Office アドインをアクティブにする Office クライアント アプリケーションを指定します。
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: c89a0154b2dbbc9b07a10493401ff761d48b955d7538eb14a825591d2b12607d
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/07/2021
ms.locfileid: "57083806"
---
# <a name="hosts-element"></a>Hosts 要素

Office アドインをアクティブにする Office クライアント アプリケーションを指定します。 **Host** 要素のコレクションとその設定が含まれます。 

[VersionOverrides](versionoverrides.md) ノードに含まれる場合、この要素は、マニフェストの親部分の **Hosts** 要素よりも優先されます。 

## <a name="child-elements"></a>子要素

|  要素 |  必須  |  説明  |
|:-----|:-----|:-----|
|  [Host](host.md)    |  はい   |  ホストとその設定について説明します。 |
