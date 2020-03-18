---
title: マニフェスト ファイルの Hosts 要素
description: Office アドインをアクティブにする Office クライアント アプリケーションを指定します。
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: cd4e0eecce610b10fdc9dafcde7b807fde425b14
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/17/2020
ms.locfileid: "42718105"
---
# <a name="hosts-element"></a>Hosts 要素

Office アドインをアクティブにする Office クライアント アプリケーションを指定します。 **Host** 要素のコレクションとその設定が含まれます。 

[VersionOverrides](versionoverrides.md) ノードに含まれる場合、この要素は、マニフェストの親部分の **Hosts** 要素よりも優先されます。 

## <a name="child-elements"></a>子要素

|  要素 |  必須  |  説明  |
|:-----|:-----|:-----|
|  [Host](host.md)    |  はい   |  ホストとその設定について説明します。 |
