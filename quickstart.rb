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

# Prints the names and majors of students in a sample spreadsheet:
# https://docs.google.com/spreadsheets/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms/edit
spreadsheet_id = "18zi9WL17Z94QdLoaCIVZLtoRXtoxj8O7bXqcfkkXu0s"

# #create columns for first name and last name
# range = "Sheet1!G1:I1"
# request_body = Google::Apis::SheetsV4::ValueRange.new
# request_body.range = range;
# request_body.values = [["First Name", "Last Name", "Velocify Status"]]
# response = service.update_spreadsheet_value spreadsheet_id, range, request_body, value_input_option: "USER_ENTERED"
# puts response
# range = "Sheet1!A2:A"
# response = service.get_spreadsheet_values spreadsheet_id, range
# puts "Name, Date:"
# puts "No data found." if response.values.empty?
# #puts response.values
# response.values.each do |row|
#   # Split full name value into two values
#   full_name = row[0].split(", ")
#   first_name = full_name[1]
#   last_name = full_name[0]
#   index = response.values.index(row)
#   rowIndex = index + 2;

  #Update columns
#   request_body2 = Google::Apis::SheetsV4::ValueRange.new
#   range2 = "Sheet1!G#{rowIndex}:H#{rowIndex}"
#   request_body2.range = range2
#   request_body2.values = [[first_name, last_name]]
#   response2 = service.update_spreadsheet_value spreadsheet_id, range2, request_body2, value_input_option: "USER_ENTERED"
# end

# Compare first and last name columns to Velocify Report

# velocify
velocify_spreadsheet_id = "11Y_L2Fkh05jZXeR5LSX3UWkgK9AfWF6vRTbQX9cegCA"
velocify_range = "Sheet1!C2:E"
velocify_response = service.get_spreadsheet_values velocify_spreadsheet_id, velocify_range

#Freedom
range = "Sheet1!G2:I"
response = service.get_spreadsheet_values spreadsheet_id, range
response.values.each do |row|
    first_name = row[0]
    last_name = row[1]
    original_status = row[2];
    index = response.values.index(row)
    rowIndex = index + 2;
    found = false
    #puts first_name + last_name
    velocify_response.values.each do |row| 
        row[0] = "" unless row[0]
        row[1] = "" unless row[1]
        if row[0].casecmp(first_name) == 0 && row[1].casecmp(last_name) == 0
            request_body2 = Google::Apis::SheetsV4::ValueRange.new
            range2 = "Sheet1!I#{rowIndex}"
            request_body2.range = range2
            status = original_status ? original_status : row[2]
            if original_status != nil
                request_body2.values = [[original_status + ", " + row[2]]]
            else
                request_body2.values = [[row[2]]]
                original_status = row[2]
            end
            response2 = service.update_spreadsheet_value spreadsheet_id, range2, request_body2, value_input_option: "USER_ENTERED"
            puts row[0] + row[1]
            found = true
            next
        end
        #puts row[0] + row[1]
    end
    if found == false
        request_body2 = Google::Apis::SheetsV4::ValueRange.new
        range2 = "Sheet1!I#{rowIndex}"
        request_body2.range = range2
        request_body2.values = [["N/A"]]
        response2 = service.update_spreadsheet_value spreadsheet_id, range2, request_body2, value_input_option: "USER_ENTERED"
    end
end
