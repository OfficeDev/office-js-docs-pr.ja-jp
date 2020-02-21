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
# <a name="create-outlook-add-ins-for-read-forms"></a><span data-ttu-id="3b8d5-103">閲覧フォーム用の Outlook アドインを作成する</span><span class="sxs-lookup"><span data-stu-id="3b8d5-103">Create Outlook add-ins for read forms</span></span>

<span data-ttu-id="3b8d5-p101">閲覧アドインは、Outlook の閲覧ウィンドウか閲覧インスペクター内でアクティブ化される Outlook アドインです。閲覧アドインは、作成アドイン (ユーザーがメッセージや予定を作成しているときにアクティブ化される Outlook アドイン) とは違って、次のユーザー シナリオで使用できます。</span><span class="sxs-lookup"><span data-stu-id="3b8d5-p101">Read add-ins are Outlook add-ins that are activated in the Reading Pane or read inspector in Outlook. Unlike compose add-ins (Outlook add-ins that are activated when a user is creating a message or appointment), read add-ins are available when users:</span></span> 

- <span data-ttu-id="3b8d5-106">電子メール メッセージ、会議出席依頼、会議の返信、または会議の取り消しの表示。</span><span class="sxs-lookup"><span data-stu-id="3b8d5-106">View an email message, meeting request, meeting response, or meeting cancellation.</span></span>

   > [!NOTE]
   > <span data-ttu-id="3b8d5-107">Outlook が閲覧フォームでアドインをアクティブ化しないメッセージの種類があります。これには、別のメッセージの添付ファイルになっているアイテムと、Outlook の [下書き] フォルダー内にあるアイテム、あるいは他の方法で暗号化または保護されているアイテムが含まれます。</span><span class="sxs-lookup"><span data-stu-id="3b8d5-107">Outlook doesn't activate add-ins in read form for certain types of messages, including items that are attachments to another message, items in the Outlook Drafts folder, or items that are encrypted or protected in other ways.</span></span>
    
- <span data-ttu-id="3b8d5-108">ユーザーが出席者になっている会議アイテムの表示。</span><span class="sxs-lookup"><span data-stu-id="3b8d5-108">View a meeting item in which the user is an attendee.</span></span>
    
- <span data-ttu-id="3b8d5-109">ユーザーが会議の開催者になっている会議アイテムの表示 (Outlook 2013 および Exchange 2013 の RTM リリースのみ)</span><span class="sxs-lookup"><span data-stu-id="3b8d5-109">View a meeting item in which the user is the organizer (RTM release of Outlook 2013 and Exchange 2013 only).</span></span>
    
   > [!NOTE]
   > <span data-ttu-id="3b8d5-p102">Office 2013 SP1 のリリースより、ユーザーが開催する会議アイテムを表示する場合、作成アドインのみをアクティブ化して使用することができます。閲覧アドインは、このシナリオでは使用できなくなります。</span><span class="sxs-lookup"><span data-stu-id="3b8d5-p102">Starting in the Office 2013 SP1 release, if the user is viewing a meeting item that the user has organized, only compose add-ins can activate and be available. Read add-ins are no longer available in this scenario.</span></span>


<span data-ttu-id="3b8d5-p103">これらの各閲覧シナリオで、アクティブ化の条件が満たされていると Outlook でアドインがアクティブ化されるので、ユーザーはアクティブ化されたアドインを閲覧ウィンドウか閲覧インスペクター内のアドイン バーで選択して開くことができます。以下の図は、ユーザーが住所を含むメッセージを閲覧するとアクティブ化されて開かれる **[Bing マップ]** アドインを示しています。</span><span class="sxs-lookup"><span data-stu-id="3b8d5-p103">In each of these read scenarios, Outlook activates add-ins when their activation conditions are fulfilled, and users can choose and open activated add-ins in the add-in bar in the Reading Pane or read inspector. The following figure shows the **Bing Maps** add-in activated and opened as the user is reading a message that contains a geographic address.</span></span>


<span data-ttu-id="3b8d5-114">**選択されている住所を含んだ Outlook メッセージに対してアクティブ化されている [Bing 地図] アドインが表示されたアドイン ウィンドウ**</span><span class="sxs-lookup"><span data-stu-id="3b8d5-114">**The add-in pane showing the Bing Maps add-in in action for the selected Outlook message that contains an address**</span></span>

![Outlook の Bing Maps メール アプリ](../images/bing-maps-add-in.jpg)


## <a name="types-of-add-ins-available-in-read-mode"></a><span data-ttu-id="3b8d5-116">閲覧モードで使用できるアドインの種類</span><span class="sxs-lookup"><span data-stu-id="3b8d5-116">Types of add-ins available in read mode</span></span>

<span data-ttu-id="3b8d5-117">閲覧アドインでは、以下のいずれの種類の組み合わせも可能です。</span><span class="sxs-lookup"><span data-stu-id="3b8d5-117">Read add-ins can be any combination of the following types.</span></span>

- [<span data-ttu-id="3b8d5-118">Outlook のアドイン コマンド</span><span class="sxs-lookup"><span data-stu-id="3b8d5-118">Add-in commands for Outlook</span></span>](add-in-commands-for-outlook.md)   
- [<span data-ttu-id="3b8d5-119">Outlook コンテキスト アドイン</span><span class="sxs-lookup"><span data-stu-id="3b8d5-119">Contextual Outlook add-ins</span></span>](contextual-outlook-add-ins.md)
    

## <a name="api-features-available-to-read-add-ins"></a><span data-ttu-id="3b8d5-120">閲覧アドインで使用できる API 機能</span><span class="sxs-lookup"><span data-stu-id="3b8d5-120">API features available to read add-ins</span></span>

- <span data-ttu-id="3b8d5-121">閲覧フォームでアドインをアクティブ化することについては、「[マニフェストでのアクティブ化ルールの指定](activation-rules.md#specify-activation-rules-in-a-manifest)」の表 1 を参照してください。</span><span class="sxs-lookup"><span data-stu-id="3b8d5-121">For activating add-ins in read forms: see Table 1 in [Specify activation rules in a manifest](activation-rules.md#specify-activation-rules-in-a-manifest).</span></span>    
- [<span data-ttu-id="3b8d5-122">正規表現アクティブ化ルールを使用して Outlook アドインを表示する</span><span class="sxs-lookup"><span data-stu-id="3b8d5-122">Use regular expression activation rules to show an Outlook add-in</span></span>](use-regular-expressions-to-show-an-outlook-add-in.md)    
- [<span data-ttu-id="3b8d5-123">Outlook アイテム内の文字列を既知のエンティティとして照合する</span><span class="sxs-lookup"><span data-stu-id="3b8d5-123">Match strings in an Outlook item as well-known entities</span></span>](match-strings-in-an-item-as-well-known-entities.md)    
- [<span data-ttu-id="3b8d5-124">Outlook アイテムからエンティティ文字列を抽出する</span><span class="sxs-lookup"><span data-stu-id="3b8d5-124">Extract entity strings from an Outlook item</span></span>](extract-entity-strings-from-an-item.md)   
- [<span data-ttu-id="3b8d5-125">サーバーから Outlook アイテムの添付ファイルを取得する</span><span class="sxs-lookup"><span data-stu-id="3b8d5-125">Get attachments of an Outlook item from the server</span></span>](get-attachments-of-an-outlook-item.md)
    

## <a name="see-also"></a><span data-ttu-id="3b8d5-126">関連項目</span><span class="sxs-lookup"><span data-stu-id="3b8d5-126">See also</span></span>

- [<span data-ttu-id="3b8d5-127">初めて Outlook アドインを記述する</span><span class="sxs-lookup"><span data-stu-id="3b8d5-127">Write your first Outlook add-in</span></span>](../quickstarts/outlook-quickstart.md)
