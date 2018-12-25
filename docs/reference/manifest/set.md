---
title: マニフェスト ファイルの Set 要素
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 0f137f7b08d6f1d0b0d972173c8085713b0f979d
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432768"
---
# <a name="set-element"></a>Set 要素

Office アドインをアクティブにするために必要な JavaScript API for Office の要件セットを指定します。

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
|名前|string|必須|[要件セット](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)の名前。|
|MinVersion|文字列|省略可能|アドインに必要な API セットの最小バージョンを指定します。**DefaultMinVersion** の値が親の [Sets](sets.md) 要素に指定されている場合は、その値を上書きします。|

## <a name="remarks"></a>解説

利用できる要件セットの詳細については、「[Office のバージョンと要件セット](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)」を参照してください。

**Set** 要素の **MinVersion** 属性と **Sets** 要素の **DefaultMinVersion** 属性の詳細については、「[マニフェストで Requirements 要素を設定する](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements#set-the-requirements-element-in-the-manifest)」をご覧ください。

> [!IMPORTANT] 
> メール アドインの場合、使用可能なのは `"Mailbox"` 要件セットのみです。 この要件セットには、Outlook のメール アドインでサポートされている API のサブセット全体が含まれ、メール アドインのマニフェストで `"Mailbox"` 要件セットを指定する必要があります (コンテンツ アドインと作業ウィンドウ アドインの場合とは異なり、オプションではありません)。 また、メール アドインで特定のメソッドのサポートを宣言することもできません。
