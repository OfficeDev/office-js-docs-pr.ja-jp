---
title: マニフェスト ファイルの Method 要素
description: Method 要素は、アクティブ化するために必要Office JavaScript API からOfficeメソッドを指定します。
ms.date: 03/19/2019
ms.localizationpriority: medium
ms.openlocfilehash: 037446f5027a97214d2b1be6ee99c8f6822b33b9
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/12/2021
ms.locfileid: "59151243"
---
# <a name="method-element"></a>Method 要素

アクティブ化するために必要Office JavaScript API Officeメソッドを指定します。

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
|名前|string|必須|必要なメソッドの名前をその親オブジェクトで修飾して指定します。 たとえば、メソッドを指定するには `getSelectedDataAsync` 、 を指定する必要があります `"Document.getSelectedDataAsync"` 。|

## <a name="remarks"></a>解説

要素 `Methods` `Method` と要素はメール アドインではサポートされていません。要件セットの詳細については、「Office[要件セット」を参照してください](../../develop/office-versions-and-requirement-sets.md)。

> [!IMPORTANT]
> 個々のメソッドの最小バージョン要件を指定する方法がないため、メソッドが実行時に使用可能であることを確認するには、そのメソッドをアドインのスクリプトで呼び出す際に、**if** ステートメントも使用する必要があります。 これを行う方法の詳細については[、「JavaScript API の概要Office参照してください](../../develop/understanding-the-javascript-api-for-office.md)。
