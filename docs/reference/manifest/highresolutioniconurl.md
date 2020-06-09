---
title: マニフェスト ファイルの HighResolutionIconUrl 要素
description: 高 DPI 画面での挿入 UX と Office ストアで Office アドインを表すために使用されるイメージの URL を指定します。
ms.date: 12/04/2018
localization_priority: Normal
ms.openlocfilehash: 1ebe69a7576c82ec121094ed84c0d49acde1098a
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611821"
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

|**属性**|**型**|**必須**|**説明**|
|:-----|:-----|:-----|:-----|
|DefaultValue|文字列 (URL)|必須|この設定の既定値を指定します。この値は、[DefaultLocale](defaultlocale.md) 要素に指定されるロケールを対象としています。|

## <a name="remarks"></a>注釈

メールアドインの場合は、[**ファイル**の管理] [アドイン] UI にアイコンが表示され  >  **Manage add-ins**ます。 コンテンツ アドインまたは作業ウィンドウ アドインでは、アイコンは、**[挿入]** > **[アドイン]** UI に表示されます。

画像のファイル形式は GIF、JPG、PNG、EXIF、BMP、TIFF のいずれかにする必要があります。 コンテンツおよび作業ウィンドウ アプリの推奨される画像の解像度は 64 x 64 ピクセルです。 メール アプリの画像は 128 × 128 ピクセルにする必要があります。 詳細については、「[効果的な AppSource と Office 内の登録リストを作成する](/office/dev/store/create-effective-office-store-listings#create-a-consistent-visual-identity)」の「_アプリに一貫性のあるビジュアル ID を作成する_」セクションを参照してください。
