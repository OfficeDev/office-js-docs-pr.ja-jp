---
title: マニフェスト ファイルの Hosts 要素
description: アドインがOfficeするクライアント アプリケーションOfficeを指定します。
ms.date: 02/25/2022
ms.localizationpriority: medium
ms.openlocfilehash: 9ea6cc9745f47b6e9b1c9bb0232b744304078053
ms.sourcegitcommit: 7b6ee73fa70b8e0ff45c68675dd26dd7a7b8c3e9
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/08/2022
ms.locfileid: "63341073"
---
# <a name="hosts-element"></a>Hosts 要素

アドインがOfficeするクライアント アプリケーションOfficeを指定します。 **Host** 要素のコレクションとその設定が含まれます。 

## <a name="as-child-of-versionoverrides-element"></a>VersionOverrides 要素の子として

このセクションの情報 *は、***Hosts** 要素が [VersionOverrides の子である場合にのみ適用されます](versionoverrides.md)。

この要素は、基本マニフェスト **の Hosts** 要素をオーバーライドします。

**アドインの種類:** 作業ウィンドウ, メール

**次の VersionOverrides スキーマでのみ有効です**。

- 作業ウィンドウ 1.0
- メール 1.0
- メール 1.1

詳細については、「Version [overrides in the manifest」を参照してください](../../develop/add-in-manifests.md#version-overrides-in-the-manifest)。

## <a name="child-elements"></a>子要素

|  要素 |  必須  |  説明  |
|:-----|:-----|:-----|
|  [Host](host.md)    |  はい   |  ホストとその設定について説明します。 |
