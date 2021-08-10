import constants
import util

# gets latests data from company facts JSON using given tag(s)
# @param json_data - company facts JSON data
# @param form_type - form type 
# @param tags - XBRL tag name
# @param units - units of given tag
# @return data in list form  
def get_latest_data_from_tags(json_data, form_type, tags, units): 
    latest_year = 0    
    data = []
    
    frame_length = 6
    if(form_type == "10-Q"): frame_length = 9

    for i in tags: 
        try: 
            current_tag_data = json_data["facts"]["us-gaap"][i]["units"][units]
            for j in range(len(current_tag_data)-1, -1, -1):
                current_year = current_tag_data[j]["fy"]
                current_form_type = current_tag_data[j]["form"]
                if(current_year == None or current_form_type == None): 
                    continue
                if("frame" in current_tag_data[j] and 
                    len(current_tag_data[j]["frame"]) != frame_length): 
                    continue
                if(int(current_year) > latest_year and current_form_type == form_type): 
                    latest_year = current_year
                    data = current_tag_data
                    break
        except KeyError as e:
            pass
    return data

# gets num_filings filing IDs from company facts JSON
# @param json_data - company facts JSON
# @param form_type - form type
# @param num_filings - number of filings requested
# @return list containing filing IDs
def get_filing_ids(json_data, form_type, num_filings):
    tags = constants.XBRL_tags["Net Income"]
    units = constants.Units_map["Net Income"]
    data = get_latest_data_from_tags(json_data, form_type, tags, units)
    
    frame_length = 6
    if(form_type == "10-Q"): frame_length = 9
    
    filing_IDs = []
    for i in range(len(data)-1, -1, -1):
        if(len(filing_IDs) == num_filings): 
            break
        if("frame" in data[i] and len(data[i]["frame"]) != frame_length): 
            continue
        if(data[i]["form"] == form_type): 
            accn_number = data[i]["accn"]
            if(accn_number not in filing_IDs): 
                filing_IDs.append(accn_number)
    return filing_IDs

# gets given quantity from company facts JSON data using latest data
# @param tags - requested tag
# @param units - units of requested data  
# @param json_data - company facts JSON data
# @param filing_ids - list containing filing IDs
# @return - data for requested tag 
def get_quantity(tags, units, json_data, form_type, filing_ids): 
    data = get_latest_data_from_tags(json_data, form_type, tags, units)   
    
    frame_length = 6
    if(form_type == "10-Q"): frame_length = 9

    values = []
    alternate_values = []
    for i in range(len(data)-1, -1, -1):
        if(len(values) == len(filing_ids)): 
            break
        if("frame" in data[i] and len(data[i]["frame"]) != frame_length and
            data[i]["form"] == form_type and data[i]["accn"] in filing_ids):
            alternate_values.append(data[i]["val"])
            continue
        if(data[i]["form"] == form_type and data[i]["accn"] in filing_ids): 
            values.append(data[i]["val"])
    
    if(values == [] and alternate_values != []): 
        values = alternate_values[:len(filing_ids)]
    elif(values == []): 
        values = [0]*len(filing_ids)
    return values

# validates whether critical inputs are loaded correctly 
# and fills in other derived/auxiliary inputs
# @param inputs - dictionary containing input values
# @param num_filings - number of filings requestes
# @return inputs (if successful), errorcode otherwise
def validate_inputs(inputs, num_filings):
    for i in range(len(inputs["Revenue"])): 
        if(inputs["Revenue"][i] == 0 or inputs["Operating Income"][i] == 0 
            or inputs["Net Income"][i] == 0):
            return constants.ERRORCODE
    if(inputs["COGS"] == [0]*num_filings):
        for i in range(num_filings):  
            cogs = (int(inputs["Revenue"][i]) - int(inputs["Operating Income"][i]))
            inputs["COGS"][i] = cogs
    if(inputs["Debt"] == [0]*num_filings):
        for i in range(num_filings): 
            debt = int(inputs["Current Long Term Debt"][i]) \
                + int(inputs["Noncurrent Long Term Debt"][i])
            inputs["Debt"][i] = debt
    return inputs

# parses required inputs from company facts JSON data given form and filings 
#@param json_data - Company facts JSON data
#@param form_type - form type
#@param num_filings - number of filings requested
def parse_inputs(json_data, form_type, num_filings):
    inputs = {}
    filing_ids = get_filing_ids(json_data, form_type, num_filings)
    assert(len(filing_ids) > 0), "Unable to find any filing IDs"

    for key in constants.XBRL_tags:
        units = constants.Units_map[key]
        tags = constants.XBRL_tags[key]
        inputs[key] = get_quantity(tags, units, json_data, form_type, filing_ids)
    inputs = validate_inputs(inputs, num_filings)
    assert(inputs != constants.ERRORCODE), "One or more of revenue, operating Income or net Income were not loaded correctly"
    return inputs

# gets inputs for given stock, form type and number of filings
# @param ticker - stock ticker on NYSE/NASDAQ
# @param form_type - form type
# @param num_filings - number of filings
# @return dictionary containing inputs 
def get_inputs(ticker, form_type="10-K", num_filings=1): 
    json_data = util.get_company_facts_json_from_ticker(ticker) 
    inputs = parse_inputs(json_data, form_type, num_filings)
    inputs["Price"] = [util.get_stock_price_from_ticker(ticker)]
    return inputs

if __name__ == "__main__":
    inputs = get_inputs("msft")
    print(inputs)