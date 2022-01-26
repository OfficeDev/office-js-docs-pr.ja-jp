---
title: マニフェスト ファイルの Method 要素
description: Method 要素は、Office アドインが Office でアクティブ化したり、基本マニフェスト設定を上書きしたりするために必要な Office JavaScript API の個々のメソッドを指定します。
ms.date: 01/22/2022
ms.localizationpriority: medium
ms.openlocfilehash: 052fb41a7077781843ea7e63d9601a819058dfa6
ms.sourcegitcommit: ae3a09d905beb4305a6ffcbc7051ad70745f79f9
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 01/26/2022
ms.locfileid: "62222270"
---
# <a name="method-element"></a>Method 要素

この要素の意味は、マニフェストで使用される場所によって異なります。

## <a name="in-the-base-manifest"></a>基本マニフェストで

基本マニフェスト (つまり、祖父母 **Requirements** 要素は [OfficeApp](officeapp.md)の直接の子) で使用する場合 **、Method** 要素は Office アドインが Office によってアクティブ化するために必要な Office JavaScript API の個別のメソッドを指定します。

**アドインの種類:** コンテンツ、作業ウィンドウ

## <a name="as-a-great-grandchild-of-a-versionoverrides-element"></a>VersionOverrides 要素の孫として

[versionOverrides](versionoverrides.md)を有効にするために、Office バージョンとプラットフォーム (Windows、Mac、Web、iOS、iPad など) でサポートする必要がある Office JavaScript API の個々のメソッドを指定します。

**アドインの種類:** 作業ウィンドウ, メール

**次の VersionOverrides スキーマでのみ有効です**。

- 祖父母の [Requirements 要素と同](requirements.md) じです。

**次の要件セットに関連付けられている**。

- 祖父母の [Requirements 要素と同](requirements.md) じです。

## <a name="syntax"></a>構文

```XML
<Method Name="string"/>
```

## <a name="contained-in"></a>含まれる場所

[Methods](methods.md)

## <a name="attributes"></a>属性

|属性|型|必須|説明|
|:-----|:-----|:-----|:-----|
|名前|string|必須|必要なメソッドの名前をその親オブジェクトで修飾して指定します。 たとえば、メソッドを指定するには `getSelectedDataAsync` 、 を指定する必要があります `"Document.getSelectedDataAsync"` 。|

## <a name="remarks"></a>注釈

基本 **マニフェストで****使用** する場合、メソッド要素とメソッド要素はメール アドインではサポートされません。 利用できる要件セットの詳細については、「[Office のバージョンと要件セット](../../develop/office-versions-and-requirement-sets.md)」を参照してください。

> [!IMPORTANT]
> 個々のメソッドの最小バージョン要件を指定する方法がないため、メソッドが実行時に使用可能であることを確認するには、そのメソッドをアドインのスクリプトで呼び出す際に、**if** ステートメントも使用する必要があります。 これを行う方法の詳細については[、「JavaScript API の概要Office参照してください](../../develop/understanding-the-javascript-api-for-office.md)。
