---
title: マニフェスト ファイルの Set 要素
description: Set 要素は、アクティブ化Officeアドインに必要Office JavaScript API 要件セットを指定します。
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: efffb3681ffb8ff310a6236c6f9aad6f001b0f86e4046df6e867798302d66f5a
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/07/2021
ms.locfileid: "57095033"
---
# <a name="set-element"></a>Set 要素

アドインがアクティブ化にOfficeする JavaScript API Officeセットを指定します。

**アドインの種類:** コンテンツ、作業ウィンドウ、メール

## <a name="syntax"></a>構文

```XML
<Set Name="string" MinVersion="n .n">
```

## <a name="contained-in"></a>含まれる場所

[Sets](sets.md)

## <a name="attributes"></a>属性

|属性|型|必須|説明|
|:-----|:-----|:-----|:-----|
|名前|string|必須|[要件セット](../../develop/office-versions-and-requirement-sets.md)の名前。|
|MinVersion|文字列|省略可能|アドインに必要な API セットの最小バージョンを指定します。 親 Sets 要素で **指定されている場合は、DefaultMinVersion** の値を [上書き](sets.md) します。|

## <a name="remarks"></a>解説

利用できる要件セットの詳細については、「[Office のバージョンと要件セット](../../develop/office-versions-and-requirement-sets.md)」を参照してください。

**Set** 要素の **MinVersion** 属性と Sets 要素 **の DefaultMinVersion** 属性の詳細については、「Set [the Requirements element in the manifest」を参照してください](../../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest)。

> [!IMPORTANT]
> メール アドインの場合、使用可能なのは `"Mailbox"` 要件セットのみです。 この要件セットには、Outlook のメール アドインでサポートされている API のサブセット全体が含まれ、メール アドインのマニフェストで `"Mailbox"` 要件セットを指定する必要があります (コンテンツ アドインと作業ウィンドウ アドインの場合とは異なり、オプションではありません)。 Also, you can't declare support for specific methods in mail add-ins.
