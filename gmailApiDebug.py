def count_emails_in_label(service, user_id, label_id):
    """
    Count the number of emails in a specific label.

    Parameters:
    - service: the authorized Gmail API service instance.
    - user_id: the user's email address or 'me' for the authenticated user.
    - label_id: the ID of the label to count emails in.

    Returns:
    - The number of emails in the specified label.
    """
    try:
        # Initialize the total count of emails
        total_emails = 0

        # Get the list of messages
        response = service.users().messages().list(userId=user_id, labelIds=label_id).execute()

        # Count the emails while there are pages of results
        while 'messages' in response:
            total_emails += len(response['messages'])
            # Check if there's another page of messages
            if 'nextPageToken' in response:
                page_token = response['nextPageToken']
                response = service.users().messages().list(userId=user_id, labelIds=label_id, pageToken=page_token).execute()
            else:
                break

        return total_emails
    except Exception as error:
        print(f"An error occurred: {error}")
        return None
    
def print_labels(service):
    results = service.users().labels().list(userId='me').execute()
    labels = results.get('labels', [])
    for label in labels:
        print(label)
