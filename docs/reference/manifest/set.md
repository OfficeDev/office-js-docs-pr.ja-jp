---
title: マニフェスト ファイルの Set 要素
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: d86b3123ff856e8618f92629308787b543e8228b
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/27/2020
ms.locfileid: "42324807"
---
# <a name="set-element"></a>Set 要素

Office アドインをアクティブにするために必要な Office JavaScript API の要件セットを指定します。

**アドインの種類:** コンテンツ、作業ウィンドウ、メール

## <a name="syntax"></a>構文

```XML
<Set Name="string" MinVersion="n .n">
```

## <a name="contained-in"></a>含まれる場所

[Sets](sets.md)

## <a name="attributes"></a>属性

|**属性**|**型**|**必須**|**説明**|
|:-----|:-----|:-----|:-----|
|名前|string|必須|[要件セット](/office/dev/add-ins/develop/office-versions-and-requirement-sets)の名前。|
|MinVersion|文字列|省略可能|アドインで必要な API セットの最小バージョンを指定します。 親[Sets](sets.md)要素で指定されている場合、 **defaultminversion**の値を上書きします。|

## <a name="remarks"></a>解説

利用できる要件セットの詳細については、「[Office のバージョンと要件セット](/office/dev/add-ins/develop/office-versions-and-requirement-sets)」を参照してください。

**Set**要素の**MinVersion**属性と**Sets**要素の**defaultminversion**属性の詳細については、「マニフェストの[要件要素を設定する](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements#set-the-requirements-element-in-the-manifest)」を参照してください。

> [!IMPORTANT] 
> メール アドインの場合、使用可能なのは `"Mailbox"` 要件セットのみです。 この要件セットには、Outlook のメール アドインでサポートされている API のサブセット全体が含まれ、メール アドインのマニフェストで `"Mailbox"` 要件セットを指定する必要があります (コンテンツ アドインと作業ウィンドウ アドインの場合とは異なり、オプションではありません)。 Also, you can't declare support for specific methods in mail add-ins.
