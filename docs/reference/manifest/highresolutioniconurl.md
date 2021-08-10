---
title: マニフェスト ファイルの HighResolutionIconUrl 要素
description: 高 DPI 画面での挿入 UX と Office ストアで Office アドインを表すために使用されるイメージの URL を指定します。
ms.date: 03/30/2021
localization_priority: Normal
ms.openlocfilehash: 4b992c7513efffe618d1b48ed89cb3b60279119c00b289a950302c9cc8e8427a
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/07/2021
ms.locfileid: "57093035"
---
# <a name="highresolutioniconurl-element"></a>HighResolutionIconUrl 要素

高 DPI 画面での挿入 UX と Office ストアで Office アドインを表すために使用されるイメージの URL を指定します。

**アドインの種類:** コンテンツ、作業ウィンドウ、メール

## <a name="syntax"></a>構文

```XML
<HighResolutionIconUrl DefaultValue="string" />
```

## <a name="can-contain"></a>含めることができるもの

[Override](override.md)

## <a name="attributes"></a>属性

|属性|型|必須|説明|
|:-----|:-----|:-----|:-----|
|DefaultValue|文字列 (URL)|必須|この設定の既定値を指定します。この値は、[DefaultLocale](defaultlocale.md) 要素に指定されるロケールを対象としています。|

## <a name="remarks"></a>注釈

メール アドインの場合、アイコンが [ファイル管理]アドイン UI  >  **に表示** されます。 コンテンツ アドインまたは作業ウィンドウ アドインでは、アイコンは、**[挿入]** > **[アドイン]** UI に表示されます。

画像のファイル形式は GIF、JPG、PNG、EXIF、BMP、TIFF のいずれかにする必要があります。 コンテンツ アプリと作業ウィンドウ アプリの場合、画像の解像度は 64 x 64 ピクセルである必要があります。 メール アプリの画像は 128 × 128 ピクセルにする必要があります。 詳細については、「[効果的な AppSource と Office 内の登録リストを作成する](/office/dev/store/create-effective-office-store-listings#create-a-consistent-visual-identity)」の「_アプリに一貫性のあるビジュアル ID を作成する_」セクションを参照してください。
