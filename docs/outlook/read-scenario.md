---
title: 閲覧フォーム用の Outlook アドインを作成する
description: 閲覧アドインは、Outlook の閲覧ウィンドウか閲覧インスペクター内でアクティブ化される Outlook アドインです。
ms.date: 04/12/2018
localization_priority: Priority
ms.openlocfilehash: a2a5448261fe6fcd150ed0cabda0184d941334e0
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/20/2020
ms.locfileid: "42166461"
---
# <a name="create-outlook-add-ins-for-read-forms"></a>閲覧フォーム用の Outlook アドインを作成する

閲覧アドインは、Outlook の閲覧ウィンドウか閲覧インスペクター内でアクティブ化される Outlook アドインです。閲覧アドインは、作成アドイン (ユーザーがメッセージや予定を作成しているときにアクティブ化される Outlook アドイン) とは違って、次のユーザー シナリオで使用できます。 

- 電子メール メッセージ、会議出席依頼、会議の返信、または会議の取り消しの表示。

   > [!NOTE]
   > Outlook が閲覧フォームでアドインをアクティブ化しないメッセージの種類があります。これには、別のメッセージの添付ファイルになっているアイテムと、Outlook の [下書き] フォルダー内にあるアイテム、あるいは他の方法で暗号化または保護されているアイテムが含まれます。
    
- ユーザーが出席者になっている会議アイテムの表示。
    
- ユーザーが会議の開催者になっている会議アイテムの表示 (Outlook 2013 および Exchange 2013 の RTM リリースのみ)
    
   > [!NOTE]
   > Office 2013 SP1 のリリースより、ユーザーが開催する会議アイテムを表示する場合、作成アドインのみをアクティブ化して使用することができます。閲覧アドインは、このシナリオでは使用できなくなります。


これらの各閲覧シナリオで、アクティブ化の条件が満たされていると Outlook でアドインがアクティブ化されるので、ユーザーはアクティブ化されたアドインを閲覧ウィンドウか閲覧インスペクター内のアドイン バーで選択して開くことができます。以下の図は、ユーザーが住所を含むメッセージを閲覧するとアクティブ化されて開かれる **[Bing マップ]** アドインを示しています。


**選択されている住所を含んだ Outlook メッセージに対してアクティブ化されている [Bing 地図] アドインが表示されたアドイン ウィンドウ**

![Outlook の Bing Maps メール アプリ](../images/bing-maps-add-in.jpg)


## <a name="types-of-add-ins-available-in-read-mode"></a>閲覧モードで使用できるアドインの種類

閲覧アドインでは、以下のいずれの種類の組み合わせも可能です。

- [Outlook のアドイン コマンド](add-in-commands-for-outlook.md)   
- [Outlook コンテキスト アドイン](contextual-outlook-add-ins.md)
    

## <a name="api-features-available-to-read-add-ins"></a>閲覧アドインで使用できる API 機能

- 閲覧フォームでアドインをアクティブ化することについては、「[マニフェストでのアクティブ化ルールの指定](activation-rules.md#specify-activation-rules-in-a-manifest)」の表 1 を参照してください。    
- [正規表現アクティブ化ルールを使用して Outlook アドインを表示する](use-regular-expressions-to-show-an-outlook-add-in.md)    
- [Outlook アイテム内の文字列を既知のエンティティとして照合する](match-strings-in-an-item-as-well-known-entities.md)    
- [Outlook アイテムからエンティティ文字列を抽出する](extract-entity-strings-from-an-item.md)   
- [サーバーから Outlook アイテムの添付ファイルを取得する](get-attachments-of-an-outlook-item.md)
    

## <a name="see-also"></a>関連項目

- [初めて Outlook アドインを記述する](../quickstarts/outlook-quickstart.md)
