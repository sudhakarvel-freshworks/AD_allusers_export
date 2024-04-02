import csv
import requests

# Function to retrieve groups for a user
def get_user_groups(user_id, access_token):
    try:
        groups_url = f"https://graph.microsoft.com/v1.0/users/{user_id}/memberOf"
        headers = {'Authorization': 'Bearer ' + access_token}
        response = requests.get(groups_url, headers=headers)
        response.raise_for_status()  # Raise an error for bad responses
        groups_data = response.json()
        return [group['displayName'] for group in groups_data.get('value', [])]
    except requests.exceptions.RequestException as e:
        print(f"Error retrieving groups for user {user_id}: {e}")
        return []
    except KeyError as ke:
        print(f"KeyError while retrieving groups for user {user_id}: {ke}")
        return []

# Function to retrieve all users in the organization
def get_all_users(access_token):
    all_users = []
    try:
        users_url = "https://graph.microsoft.com/v1.0/users"
        headers = {'Authorization': 'Bearer ' + access_token}
        while users_url:
            response = requests.get(users_url, headers=headers)
            response.raise_for_status()  # Raise an error for bad responses
            users_data = response.json()
            all_users.extend(users_data.get('value', []))
            users_url = users_data.get('@odata.nextLink')
        return all_users
    except requests.exceptions.RequestException as e:
        print(f"Error retrieving users: {e}")
        return []

# Main function
def main():
    # Replace with your access token
    access_token = '<token>'
    try:
        # Retrieve all users in the organization
        print("Retrieving all users...")
        all_users = get_all_users(access_token)
        print(f"Total users retrieved: {len(all_users)}")

        # Write user details and their groups to CSV file
        with open('user_groups.csv', mode='w', newline='', encoding='utf-8') as file:
            writer = csv.writer(file)
            writer.writerow(['User Name', 'Email', 'User ID', 'Groups', 'Active State'])

            for user in all_users:
                user_id = user.get('id', '')
                user_name = user.get('displayName', '')
                user_email = user.get('mail', '')
                user_active = "True" if user.get('accountEnabled', True) else "False"
                print(f"Processing user: {user_name}")
                user_groups = get_user_groups(user_id, access_token)
                writer.writerow([user_name, user_email, user_id, ', '.join(user_groups), user_active])
                print(f"{user_name} processed.")
        print("User details exported to user_groups.csv")
    except requests.exceptions.RequestException as e:
        print(f"An error occurred while making a request: {e}")
    except IOError as ioe:
        print(f"An I/O error occurred: {ioe}")
    except Exception as ex:
        print(f"An unexpected error occurred: {ex}")

if __name__ == "__main__":
    main()
