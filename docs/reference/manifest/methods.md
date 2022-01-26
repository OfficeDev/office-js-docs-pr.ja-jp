---
title: マニフェスト ファイルの Methods 要素
description: Methods 要素は、Office でアクティブ化したり、基本マニフェスト設定を上書きしたりするために、Office アドインが必要とする Office JavaScript API メソッドの一覧を指定します。
ms.date: 01/22/2022
ms.localizationpriority: medium
ms.openlocfilehash: 4c39c6363cd33e103cf40c0f7f047fa694db1411
ms.sourcegitcommit: ae3a09d905beb4305a6ffcbc7051ad70745f79f9
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 01/26/2022
ms.locfileid: "62222277"
---
# <a name="methods-element"></a>Methods 要素

この要素の意味は、マニフェストで使用される場所によって異なります。

## <a name="in-the-base-manifest"></a>基本マニフェストで

基本マニフェストで使用する場合 (つまり、親 **Requirements** 要素は [OfficeApp](officeapp.md)の直接の子です **)、Methods** 要素は、Office でアクティブ化するために Office アドインが必要とする Office JavaScript API メソッドの一覧を指定します。

**アドインの種類:** コンテンツ、作業ウィンドウ

## <a name="as-a-grandchild-of-a-versionoverrides-element"></a>VersionOverrides 要素の孫として

[VersionOverrides](versionoverrides.md)を有効にするために、Office バージョンとプラットフォーム (Windows、Mac、Web、iOS、iPad など) でサポートする必要がある Office JavaScript API メソッドの最小セットを指定します。

**アドインの種類:** 作業ウィンドウ, メール

**次の VersionOverrides スキーマでのみ有効です**。

- 親の [Requirements 要素と同](requirements.md) じです。

**次の要件セットに関連付けられている**。

- 親の [Requirements 要素と同](requirements.md) じです。

## <a name="syntax"></a>構文

```XML
<Methods>
   ...
</Methods>
```

## <a name="contained-in"></a>含まれる場所

[Requirements](requirements.md)

## <a name="can-contain"></a>含めることができるもの

[Method](method.md)

## <a name="remarks"></a>注釈

基本 **マニフェストで****使用** する場合、メソッド要素とメソッド要素はメール アドインではサポートされません。 利用できる要件セットの詳細については、「[Office のバージョンと要件セット](../../develop/office-versions-and-requirement-sets.md)」を参照してください。
