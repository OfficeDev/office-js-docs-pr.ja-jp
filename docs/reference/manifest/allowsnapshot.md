---
title: マニフェスト ファイルの AllowSnapshot 要素
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: f1aced0ce37b01c277ea5a8621f6c7764d2f761b
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432348"
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
 > **AllowSnapshot** の既定値は `true` です。 この場合、Office アドインをサポートしていないバージョンのホスト アプリケーションでドキュメントを開くユーザーがアドインのイメージを表示できるようになったり、ホスト アプリケーションがアドインをホストするサーバーに接続できない場合にアドインの静的イメージが提供されたりします。 しかしこれは、アドインがホストされるドキュメントから、アドインに表示される機密性の高い情報に直接アクセスできるということでもあります。

