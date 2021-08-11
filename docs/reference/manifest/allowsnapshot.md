---
title: マニフェスト ファイルの AllowSnapshot 要素
description: ホスト ドキュメントと共にコンテンツ アドインのスナップショット イメージを保存するかどうかを指定します。
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 1462b60dffda7e3bb611225f015b5a1c9f0b5e78271580383961cc118af60587
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/07/2021
ms.locfileid: "57095057"
---
# <a name="allowsnapshot-element"></a>AllowSnapshot 要素

ホスト ドキュメントと共にコンテンツ アドインのスナップショット イメージを保存するかどうかを指定します。

**アドインの種類:** コンテンツ

## <a name="syntax"></a>構文

```XML
<AllowSnapshot> [true | false]</AllowSnapshot>
```

## <a name="contained-in"></a>含まれる場所

[OfficeApp](officeapp.md)

## <a name="remarks"></a>解説

 > [!IMPORTANT]
 > **AllowSnapshot** の既定値は `true` です。 これにより、Office Office アドインをサポートしないバージョンの Office アプリケーションでドキュメントを開くユーザーに対してアドインのイメージが表示され、アプリケーションがアドインをホストするサーバーに接続できない場合は、アドインの静的イメージが提供されます。 However, this also means that potentially sensitive information displayed in the add-in can be accessed directly from the document hosting the add-in.
