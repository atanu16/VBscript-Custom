import sys
from datetime import datetime
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.user_credential import UserCredential

# ----------------------
# Get input arguments
# ----------------------
bot_name = sys.argv[1]
platform = sys.argv[2]
username_val = sys.argv[3]
machine = sys.argv[4]
status = sys.argv[5]
region = sys.argv[6]
start_time = sys.argv[7]
end_time = sys.argv[8]

# ----------------------
# SharePoint Config
# ----------------------
site_url = "https://yourtenant.sharepoint.com/sites/Dev16"
list_name = "Temp"
username = "your_email@example.com"
password = "your_password"

def get_list_item(ctx, bot_name):
    caml = f"""
    <View>
        <Query>
            <Where>
                <Eq>
                    <FieldRef Name='Title' />
                    <Value Type='Text'>{bot_name}</Value>
                </Eq>
            </Where>
        </Query>
    </View>
    """
    items = ctx.web.lists.get_by_title(list_name).get_items_from_xml(caml)
    ctx.load(items)
    ctx.execute_query()
    return items

def create_new_item(ctx):
    item = ctx.web.lists.get_by_title(list_name).add_item({
        'Title': bot_name,
        'Platform': platform,
        'Username': username_val,
        'Machine': machine,
        'Date': datetime.now().strftime('%Y-%m-%d'),
        'StartTime': start_time,
        'EndTime': end_time,
        'Status': status,
        'Region': region
    })
    ctx.execute_query()
    print("New item created.")

def update_item(ctx, item):
    item.set_property('Platform', platform)
    item.set_property('Username', username_val)
    item.set_property('Machine', machine)
    item.set_property('StartTime', start_time)
    item.set_property('EndTime', end_time)
    item.set_property('Status', status)
    item.set_property('Region', region)
    item.update()
    ctx.execute_query()
    print("Item updated.")

# Main execution
ctx = ClientContext(site_url).with_credentials(UserCredential(username, password))
items = get_list_item(ctx, bot_name)

if items:
    if status.lower() == "completed":
        create_new_item(ctx)
    else:
        update_item(ctx, items[0])
else:
    create_new_item(ctx)