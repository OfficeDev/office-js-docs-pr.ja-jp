---
title: マニフェスト ファイルの Hosts 要素
description: Office アドインをアクティブにする Office クライアント アプリケーションを指定します。
ms.date: 10/09/2018
ms.localizationpriority: medium
ms.openlocfilehash: 2684753fc32a295d7e177ef3bf668c194458128e
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/12/2021
ms.locfileid: "59151446"
---
# <a name="hosts-element"></a>Hosts 要素

Office アドインをアクティブにする Office クライアント アプリケーションを指定します。 **Host** 要素のコレクションとその設定が含まれます。 

[VersionOverrides](versionoverrides.md) ノードに含まれる場合、この要素は、マニフェストの親部分の **Hosts** 要素よりも優先されます。 

## <a name="child-elements"></a>子要素

|  要素 |  必須  |  説明  |
|:-----|:-----|:-----|
|  [Host](host.md)    |  はい   |  ホストとその設定について説明します。 |
