---
title: マニフェスト ファイルの Set 要素
description: Set 要素は、Office によってアクティブ化したり、基本マニフェスト設定を上書きしたりするために必要な Office Office アドインに必要な JavaScript API 要件セットを指定します。
ms.date: 01/22/2022
ms.localizationpriority: medium
ms.openlocfilehash: 55e1b25765bfbe53108bc9201c0c851c6ef9161d
ms.sourcegitcommit: ae3a09d905beb4305a6ffcbc7051ad70745f79f9
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 01/26/2022
ms.locfileid: "62222235"
---
# <a name="set-element"></a>Set 要素

この要素の意味は、マニフェストで使用される場所によって異なります。

## <a name="in-the-base-manifest"></a>基本マニフェストで

基本マニフェスト (つまり、祖父母 **Requirements** 要素は [OfficeApp](officeapp.md)の直接の子) で使用する場合 [](../../develop/office-versions-and-requirement-sets.md#specify-office-applications-and-requirement-sets)**、Set** 要素は Office アドインが Office によってアクティブ化するために必要な Office JavaScript API からの要件セットを指定します。

**アドインの種類:** コンテンツ、作業ウィンドウ、メール

## <a name="as-a-great-grandchild-of-a-versionoverrides-element"></a>VersionOverrides 要素の孫として

[versionOverrides](versionoverrides.md)を有効にするために、Office バージョンとプラットフォーム (Windows、Mac、Web、iOS、iPad など) でサポートする必要がある Office JavaScript API の要件セットを指定します。 [](../../develop/office-versions-and-requirement-sets.md#specify-office-applications-and-requirement-sets)

**アドインの種類:** 作業ウィンドウ, メール

**次の VersionOverrides スキーマでのみ有効です**。

- 祖父母の [Requirements 要素と同](requirements.md) じです。

**次の要件セットに関連付けられている**。

- 祖父母の [Requirements 要素と同](requirements.md) じです。

## <a name="syntax"></a>構文

```XML
<Set Name="string" MinVersion="n .n">
```

## <a name="contained-in"></a>含まれる場所

[Sets](sets.md)

## <a name="attributes"></a>属性

|属性|種類|必須|説明|
|:-----|:-----|:-----|:-----|
|名前|string|必須|[要件セット](../../develop/office-versions-and-requirement-sets.md)の名前。|
|MinVersion|文字列|省略可能|アドインに必要な API セットの最小バージョンを指定します。 親 Sets 要素で **指定されている場合は、DefaultMinVersion** の値を [上書き](sets.md) します。|

## <a name="remarks"></a>解説

利用できる要件セットの詳細については、「[Office のバージョンと要件セット](../../develop/office-versions-and-requirement-sets.md)」を参照してください。

**Set** 要素の **MinVersion** 属性と **Sets** 要素の **DefaultMinVersion** 属性の詳細については、「アドインをホストできる Office バージョンとプラットフォームを指定する」を [参照してください](../../develop/specify-office-hosts-and-api-requirements.md#specify-which-office-versions-and-platforms-can-host-your-add-in)。

