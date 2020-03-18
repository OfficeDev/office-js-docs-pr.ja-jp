---
title: マニフェスト ファイルの AllowSnapshot 要素
description: ホスト ドキュメントと共にコンテンツ アドインのスナップショット イメージを保存するかどうかを指定します。
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 8bb143d13a17b3e184af64f1bf18f2a32a55b60c
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720961"
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
 > **AllowSnapshot** の既定値は `true` です。 この場合、Office アドインをサポートしていないバージョンのホスト アプリケーションでドキュメントを開くユーザーがアドインのイメージを表示できるようになったり、ホスト アプリケーションがアドインをホストするサーバーに接続できない場合にアドインの静的イメージが提供されたりします。 However, this also means that potentially sensitive information displayed in the add-in can be accessed directly from the document hosting the add-in.

