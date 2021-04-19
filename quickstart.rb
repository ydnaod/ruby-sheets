require "google/apis/sheets_v4"
require "googleauth"
require "googleauth/stores/file_token_store"
require "fileutils"

OOB_URI = "urn:ietf:wg:oauth:2.0:oob".freeze
APPLICATION_NAME = "Google Sheets API Ruby Quickstart".freeze
CREDENTIALS_PATH = "credentials.json".freeze
# The file token.yaml stores the user's access and refresh tokens, and is
# created automatically when the authorization flow completes for the first
# time.
TOKEN_PATH = "token.yaml".freeze
SCOPE = Google::Apis::SheetsV4::AUTH_SPREADSHEETS

##
# Ensure valid credentials, either by restoring from the saved credentials
# files or intitiating an OAuth2 authorization. If authorization is required,
# the user's default browser will be launched to approve the request.
#
# @return [Google::Auth::UserRefreshCredentials] OAuth2 credentials
def authorize
  client_id = Google::Auth::ClientId.from_file CREDENTIALS_PATH
  token_store = Google::Auth::Stores::FileTokenStore.new file: TOKEN_PATH
  authorizer = Google::Auth::UserAuthorizer.new client_id, SCOPE, token_store
  user_id = "default"
  credentials = authorizer.get_credentials user_id
  if credentials.nil?
    url = authorizer.get_authorization_url base_url: OOB_URI
    puts "Open the following URL in the browser and enter the " \
         "resulting code after authorization:\n" + url
    code = gets
    credentials = authorizer.get_and_store_credentials_from_code(
      user_id: user_id, code: code, base_url: OOB_URI
    )
  end
  credentials
end

# Initialize the API
service = Google::Apis::SheetsV4::SheetsService.new
service.client_options.application_name = APPLICATION_NAME
service.authorization = authorize

# Stroudsburg
spreadsheet_id = "18zi9WL17Z94QdLoaCIVZLtoRXtoxj8O7bXqcfkkXu0s"

def write_values (range, values, service, spreadsheet_id)
    request_body2 = Google::Apis::SheetsV4::ValueRange.new
    request_body2.range = range
    request_body2.values = values
    # puts values.inspect
    response2 = service.update_spreadsheet_value spreadsheet_id, range, request_body2, value_input_option: "USER_ENTERED"
end

#create columns for first name and last name
range = "Sheet1!G1:K1"
request_body = Google::Apis::SheetsV4::ValueRange.new
request_body.range = range;
request_body.values = [["First Name", "Last Name", "Velocify Status", "", "Lead Source"]]
response = service.update_spreadsheet_value spreadsheet_id, range, request_body, value_input_option: "USER_ENTERED"
# puts response
range = "Sheet1!A2:A"
response = service.get_spreadsheet_values spreadsheet_id, range
puts "No data found." if response.values.empty?
#puts response.values
values = []
response.values.each do |row|
  # Split full name value into two values
  full_name = row[0].split(", ")
  first_name = full_name[1]
  last_name = full_name[0]
  index = response.values.index(row)
  rowIndex = index + 2;

  #Update columns
    name_array = []
    name_array.push(first_name)
    name_array.push(last_name)
    values.push(name_array)
end
# puts values.inspect
write_values("Sheet1!G2:H", values, service, spreadsheet_id)


# Compare first and last name columns to Velocify Report

# velocify
velocify_spreadsheet_id = "1_XiOlMEypPgXXVw0VgEA-0veF4t73qRkjM7bgRX26wE"
velocify_range = "Sheet1!C2:H"
velocify_response = service.get_spreadsheet_values velocify_spreadsheet_id, velocify_range

#Freedom
range = "Sheet1!G2:K"
response = service.get_spreadsheet_values spreadsheet_id, range
values = []
lead_sources = [];
response.values.each do |row|
    first_name = row[0]
    last_name = row[1]
    original_status = row[2];

    original_lead_source = row[4];
    index = response.values.index(row)
    rowIndex = index + 2;
    found = false
    #puts first_name + last_name
    velocify_response.values.each do |row| 
        row[0] = "" unless row[0]
        row[1] = "" unless row[1]
        if row[0].casecmp(first_name) == 0 && row[1].casecmp(last_name) == 0
            status = original_status ? original_status : row[2]

            if original_status != nil
                temp_string = original_status + ", " + row[2]
                values[index] = [temp_string]
                original_status = temp_string
            else
                values[index] = [row[2]]
                original_status = row[2]
            end

            if original_lead_source != nil
                temp_string = original_lead_source + ', ' + row[5]
                lead_sources[index] = [temp_string]
                original_lead_source = temp_string
            else
                lead_sources[index] = [row[5]]
                original_lead_source = row[5]
            end
            
            found = true
            next
        end
    end
    if found == false
        values[index] = ["N/A"]
        lead_sources[index] = ["N/A"]
    end
end
puts lead_sources.inspect
write_values("Sheet1!I2:I", values, service, spreadsheet_id)
write_values("Sheet1!K2:K", lead_sources, service, spreadsheet_id)
