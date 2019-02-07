require 'json'
require 'net/http'
require 'openssl'

require 'google/apis/sheets_v4'
require 'googleauth'
require 'googleauth/stores/file_token_store'
require 'fileutils'

OOB_URI = 'urn:ietf:wg:oauth:2.0:oob'.freeze
APPLICATION_NAME = 'Vizex Supplies'.freeze
CREDENTIALS_PATH = 'credentials.json'.freeze
# The file token.yaml stores the user's access and refresh tokens, and is
# created automatically when the authorization flow completes for the first time.
TOKEN_PATH = 'token.yaml'.freeze
SCOPE = Google::Apis::SheetsV4::AUTH_SPREADSHEETS

##
# Ensure valid credentials, either by restoring from the saved credentials
# files or intitiating an OAuth2 authorization. If authorization is required,
# the user's default browser will be launched to approve the request.
# @return [Google::Auth::UserRefreshCredentials] OAuth2 credentials
def authorize
	client_id = Google::Auth::ClientId.from_file(CREDENTIALS_PATH)
	token_store = Google::Auth::Stores::FileTokenStore.new(file: TOKEN_PATH)
	authorizer = Google::Auth::UserAuthorizer.new(client_id, SCOPE, token_store)
	user_id = 'default'
	credentials = authorizer.get_credentials(user_id)
	if credentials.nil?
	url = authorizer.get_authorization_url(base_url: OOB_URI)
	puts 'Open the following URL in the browser and enter the ' \
	     "resulting code after authorization:\n" + url
	code = gets
	credentials = authorizer.get_and_store_credentials_from_code(
		user_id: user_id, code: code, base_url: OOB_URI
	)
	end
	credentials
end

# Use Trello API documentation instructions to get your own Key and token.
url = 'https://api.trello.com/1/boards/bsjhgc2U/cards?fields=name,desc,idMembersVoted,idList&key=yourkey&token=yourtoken'
uri = URI(url)

http = Net::HTTP.new(uri.host, uri.port)
http.use_ssl = true
http.verify_mode = OpenSSL::SSL::VERIFY_NONE

request = Net::HTTP::Get.new(url)

response = http.request(request)
parsed = JSON.parse(response.read_body)

values = []
parsed.each do |elem|
	values.push( [0, elem['name'], elem['desc'], elem['idMembersVoted'].length ] ) if elem['idList'] == '5c3f0ab4d18b616162e85ef2'
end

values = values.sort do |a,b|
	b[3] <=> a[3]
end

counter = 0
values.each do |elem|
	elem [0] = counter += 1
end

# Initialize the API.
service = Google::Apis::SheetsV4::SheetsService.new
service.client_options.application_name = APPLICATION_NAME
service.authorization = authorize

range_name = 'A3'
value_range_object = Google::Apis::SheetsV4::ValueRange.new(range: range_name, values: values)

spreadsheet_id = '1bKAArbLZevkexeWMMoX1Um2wBCUCgjrIiIdrIWKE47U'
result = service.update_spreadsheet_value(spreadsheet_id,
                                          range_name,
                                          value_range_object,
                                          value_input_option: 'RAW')

puts "#{result.updated_cells} cells updated."
