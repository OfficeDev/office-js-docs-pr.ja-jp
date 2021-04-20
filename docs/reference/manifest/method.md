---
title: マニフェスト ファイルの Method 要素
description: Method 要素は、office アドインをアクティブにするために必要な、Office JavaScript API からの個別のメソッドを指定します。
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 0e3e74a73a3422a7789e82d6f0e7a516bd795ca8
ms.sourcegitcommit: cc6886b47c84ac37a3c957ff85dd0ed526ca5e43
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/12/2020
ms.locfileid: "46641326"
---
# <a name="method-element"></a>Method 要素

Office JavaScript API から、Office アドインをアクティブにするために必要な個別のメソッドを指定します。

**アドインの種類:** コンテンツ、作業ウィンドウ

## <a name="syntax"></a>構文

```XML
<Method Name="string"/>
```

## <a name="contained-in"></a>含まれる場所

[Methods](methods.md)

## <a name="attributes"></a>属性

|属性|型|必須|説明|
|:-----|:-----|:-----|:-----|
|名前|string|必須|必要なメソッドの名前をその親オブジェクトで修飾して指定します。 たとえば、メソッドを指定するには、を `getSelectedDataAsync` 指定する必要があり `"Document.getSelectedDataAsync"` ます。|

## <a name="remarks"></a>注釈

`Methods`および要素は、 `Method` メールアドインではサポートされていません。要件セットの詳細については、「 [Office のバージョンと要件セット](../../develop/office-versions-and-requirement-sets.md)」を参照してください。

> [!IMPORTANT]
> 個々のメソッドの最小バージョン要件を指定する方法がないため、メソッドが実行時に使用可能であることを確認するには、そのメソッドをアドインのスクリプトで呼び出す際に、**if** ステートメントも使用する必要があります。 これを行う方法の詳細については、「 [Office JAVASCRIPT API に](../../develop/understanding-the-javascript-api-for-office.md)ついて」を参照してください。
