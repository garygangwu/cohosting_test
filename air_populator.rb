require "google/api_client"
require "google_drive"
require 'net/http'
require 'date'
require './airconst'

class GoogleOAuth2Utils
  
  # /usr/local/share/ruby/gems/2.0/gems/google_drive-1.0.1/lib/google_drive

  def self.get_access_token
    client = Google::APIClient.new(
      :application_name => 'Cohosting Experimental App',
      :application_version => '1.0.0')
    client.authorization.client_id = AirConst::CLIENT_ID
    client.authorization.client_secret = AirConst::CLIENT_SECRET
    client.authorization.grant_type = 'refresh_token'
    client.authorization.refresh_token = AirConst::REFRESH_TOKEN

    client.authorization.fetch_access_token!
    return client.authorization.access_token
  end

  def self.get_work_sheets
    access_token = self.get_access_token
    session = GoogleDrive.login_with_oauth(access_token)
    return session.spreadsheet_by_key(AirConst::SHEET_KEY).worksheets
  end
end

class AirbnbUtils
  def self.get_all_reservations(user_id)
    access_token = AirConst::USER_TOKENS[user_id]
    if access_token.nil?
      return nil
    end

    uri = URI('https://api.airbnb.com/v1/reservations?role=host&items_per_page=100')   
    site = create_endpoint(uri)
    headers = {}
    headers['Accept'] = 'application/json'
    headers['Content-Type'] = 'application/json'
    headers['X-Airbnb-OAuth-Token'] = access_token
    response = site.send_request('GET', uri.request_uri, data = nil, headers)
    body = response.body.strip
    result = body.empty? ? {} : JSON.parse(body)
    return result
  end

  private

  def self.create_endpoint(uri)
    site = Net::HTTP.new(uri.host, uri.port)
    site.read_timeout = 20
    if uri.scheme.downcase == 'https'
      site.use_ssl = true
      site.verify_mode = OpenSSL::SSL::VERIFY_NONE
    end
    return site
  end
end

class DataUtils
  def self.convert_airbnb_reservations_to_list(reservation_json_obj)
    reservation_list = {}
    if reservation_json_obj.nil? || reservation_json_obj['reservations'].nil?
      return reservation_list
    end
    reservation_json_obj['reservations'].each do |value|
      reservation = value['reservation']
      confirmation_code = reservation['confirmation_code']
      r = {}
      r['confirmation_code'] = confirmation_code
      r['check_in'] = reservation['start_date']
      date = Date.parse(r['check_in'])
      r['check_out'] = (date + reservation['nights']).to_s
      r['guest'] = reservation['guest']['user']['first_name'] + ' ' + reservation['guest']['user']['last_name']
      r['number_of_guests'] = reservation['number_of_guests']
      r['guest_email'] = reservation['guest']['user']['email']
      r['host'] = reservation['host']['user']['first_name']
      r['listing'] = reservation['listing']['listing']['name']
      reservation_list[confirmation_code] = r
    end
    return reservation_list
  end

  def self.convert_work_sheet_to_list(rows)
    {}
  end

  def self.merge_into_sorted_arrays(airbnb, sheet)
    reservations = []
    airbnb.each do |k, v|
      reservations << v
    end
    sheet.each do |k, v|
      if airbnb[k].nil?
        reservations << v
      end
    end
    reservations.sort! { |a, b| a['check_in'] <=> b['check_in'] }
  end

  #{"confirmation_code"=>"5BN8QR", "check_in"=>"2015-06-30", "check_out"=>"2015-07-05", "guest"=>"Toshihiro Osada", "number_of_guests"=>2, "guest_email"=>"toshihiro-qsbndbzd6pjn3v9v@guest.airbnb.com", "host"=>"Maíra", "listing"=>"Perfect size studio in SOMA"}
  def self.convert_to_sheet_rows(reservations)
    rows = []
    reservations.each do |r|
      row = []
      row << r['confirmation_code']
      row << r['host']
      row << r['listing']
      row << r['check_in']
      row << r['check_out']
      row << r['guest']
      row << r['number_of_guests']
      row << r['guest_email']
      rows << row
    end
    return rows
  end
end

reservation_json_obj = AirbnbUtils.get_all_reservations(2529228)
airbnb_list = DataUtils.convert_airbnb_reservations_to_list(reservation_json_obj)
ws = GoogleOAuth2Utils::get_work_sheets()[0]
sheet_list = DataUtils.convert_work_sheet_to_list(ws.rows)
reservations = DataUtils.merge_into_sorted_arrays(airbnb_list, sheet_list)
rows = DataUtils.convert_to_sheet_rows(reservations)
ws.update_cells(2, 1, rows)
ws.save
