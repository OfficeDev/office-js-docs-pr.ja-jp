---
title: マニフェスト ファイルの Version 要素
description: Version 要素は、アドイン Officeを指定します。
ms.date: 02/05/2021
ms.localizationpriority: medium
ms.openlocfilehash: 34cefa22123ed4ee723d51a669e01e042efc2934
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/12/2021
ms.locfileid: "59154684"
---
# <a name="version-element"></a>Version 要素

Office アドインのバージョンを指定します。 バージョン番号は、1、2、3、または 4 パーツ (n、n.n、n.n.n、または n.n.n.n) です。

**アドインの種類:** コンテンツ、作業ウィンドウ、メール

## <a name="syntax"></a>構文

```XML
<Version>n[.n.n.n]</Version>
```

## <a name="contained-in"></a>含まれる場所

[OfficeApp](officeapp.md)

## <a name="remarks"></a>注釈

バージョン番号の各部分には、最大 5 桁の数字を指定できます。
