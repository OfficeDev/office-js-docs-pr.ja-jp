---
title: マニフェスト ファイルの IconUrl 要素
description: IconUrl 要素は、挿入 UX と Office ストアで Office アドインを表すイメージの URL を指定します。
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: b2c117f69259911bf4643eaddbbbcdf5fb502197
ms.sourcegitcommit: cc6886b47c84ac37a3c957ff85dd0ed526ca5e43
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/12/2020
ms.locfileid: "46641362"
---
# <a name="iconurl-element"></a>IconUrl 要素

挿入 UX と Office ストアで Office アドインを表すために使用されるイメージの URL を指定します。

**アドインの種類:** コンテンツ、作業ウィンドウ、メール

## <a name="syntax"></a>構文

```XML
<IconUrl DefaultValue="string" />
```

## <a name="can-contain"></a>含めることができるもの

[Override](override.md)

## <a name="attributes"></a>属性

|属性|型|必須|説明|
|:-----|:-----|:-----|:-----|
|DefaultValue|文字列|必須|この設定の既定値を指定します。この値は、[DefaultLocale](defaultlocale.md) 要素に指定されるロケールを対象としています。|

## <a name="remarks"></a>注釈

メールアドインの場合は、[**ファイル**の  >  **管理**] ui (outlook) または [設定] [アドインの**Settings**  >  **管理**] ui (web 上の outlook) にアイコンが表示されます。 コンテンツ アドインまたは作業ウィンドウ アドインでは、アイコンは、**[挿入]** > **[アドイン]** UI に表示されます。 すべてのアドインの種類について、アドインを AppSource に発行する場合、このアイコンは[appsource](https://appsource.microsoft.com)でも使用されます。

画像のファイル形式は GIF、JPG、PNG、EXIF、BMP、TIFF のいずれかにする必要があります。 コンテンツ アプリおよび作業ウィンドウ アプリの場合、指定する画像は 32 x 32 ピクセルにする必要があります。 メール アプリの場合、推奨される画像の解像度は 64 x 64 ピクセルです。 [HighResolutionIconUrl](highresolutioniconurl.md) 要素を使用して、高 DPI 画面で実行されている Office ホスト アプリケーションで使用するアイコンも指定する必要があります。 詳細については、「[効果的な AppSource と Office 内の登録リストを作成する](/office/dev/store/create-effective-office-store-listings#create-a-consistent-visual-identity)」の「_アプリに一貫性のあるビジュアル ID を作成する_」セクションを参照してください。

実行時に要素の値を変更する `IconUrl` ことは現在サポートされていません。