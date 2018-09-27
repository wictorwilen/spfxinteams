const eventTemplate: any = {
    "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
    "type": "AdaptiveCard",
    "version": "1.0",
    "body": [
        {
            "type": "TextBlock",
            "text": "TITLE",
            "size": "large",
            "weight": "bolder",
            "color": "accent"
        },
        {
            "type": "TextBlock",
            "text": "CONFERENCE ROOM",
            "isSubtle": true
        },
        {
            "type": "TextBlock",
            "text": "TIME",
            "isSubtle": true,
            "spacing": "none"
        }
    ],
    "actions": [
        {
            "type": "Action.Submit",
            "title": "Add to Outlook",
            "id": "outlook",
            "data": {
                "x": "addToOutlook"
            }
        },
        {
            "type": "Action.ShowCard",
            "title": "Comment",
            "card": {
                "type": "AdaptiveCard",
                "body": [
                    {
                        "type": "Container",
                        "items": [
                            {
                                "type": "TextBlock",
                                "text": "Channel",
                                "isSubtle": true,
                                "spacing": "none"
                            },
                            {
                                "type": "Input.ChoiceSet",
                                "id": "channel",
                                "style": "compact",
                                "choices": [
                                    {
                                        "title": "General",
                                        "value": "General",
                                        "isSelected": true
                                    }
                                ]
                            }]
                    },
                    {
                        "type": "Input.Text",
                        "id": "comment",
                        "isMultiline": true,
                        "placeholder": "Enter your comment"
                    }
                ],
                "actions": [
                    {
                        "id": "comment",
                        "type": "Action.Submit",
                        "title": "OK"
                    }
                ]
            }
        }
    ]
};

export default eventTemplate;