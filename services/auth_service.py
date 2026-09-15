from .excel_service import read_sheet

USERS_SHEET = "Users Data"


def verify_user(user_id, password):
    """Verify user credentials and return user details if valid."""

    df = read_sheet(USERS_SHEET)

    # Debug: Check the columns loaded from Excel
    print("COLUMNS:", df.columns.tolist())

    # Find the user
    row = df.loc[df["Users ID"].astype(str) == str(user_id)]

    # Debug: Check whether the user was found
    print("USER ID ENTERED:", user_id)
    print("USER FOUND:", not row.empty)

    if row.empty:
        print("DEBUG: User not found.")
        return None

    user_record = row.iloc[0]

    # Debug: Show user information (do NOT print password)
    print("DEBUG: User record found.")
    print("NAME:", user_record["Name"])
    print("ROLE:", user_record["Role"])

    stored_password = str(user_record["Password"])

    # Debug: Check password result
    if stored_password == str(password):
        print("DEBUG: Password verified successfully.")

        return {
            "user_id": user_id,
            "name": user_record["Name"],
            "role": user_record["Role"],
        }

    print("DEBUG: Incorrect password.")
    return None