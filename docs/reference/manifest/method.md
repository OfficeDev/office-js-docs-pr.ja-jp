---
title: マニフェスト ファイルの Method 要素
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: fded84344182bb45597b00a794f18defaa44d3b3
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432824"
---
# <a name="method-element"></a>Method 要素

Office アドインをアクティブにするために必要な JavaScript API for Office の個別のメソッドを指定します。

**アドインの種類:** コンテンツ、作業ウィンドウ

## <a name="syntax"></a>構文

```XML
<Method Name="string"/>
```

## <a name="contained-in"></a>含まれる場所

[Methods](methods.md)

## <a name="attributes"></a>属性

|**属性**|**型**|**必須**|**説明**|
|:-----|:-----|:-----|:-----|
|名前|string|必須|必要なメソッドの名前をその親オブジェクトで修飾して指定します。たとえば、**getSelectedDataAsync** メソッドを指定するには、`"Document.getSelectedDataAsync"` と指定する必要があります。|

## <a name="remarks"></a>解説

**Methods** と **Method** 要素はメール アドインではサポートされていません。要件セットの詳細については、「[Office のバージョンと要件セット](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)」をご覧ください。

> [!IMPORTANT] 
> 個々のメソッドの最小バージョン要件を指定する方法がないため、メソッドが実行時に使用可能であることを確認するには、そのメソッドをアドインのスクリプトで呼び出す際に、**if** ステートメントも使用する必要があります。 これを行う方法の詳細については、「[JavaScript API for Office について](https://docs.microsoft.com/office/dev/add-ins/develop/understanding-the-javascript-api-for-office)」をご覧ください。

