import constants
import util

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
                if("frame" in current_tag_data[j] and len(current_tag_data[j]["frame"]) != frame_length): 
                    continue
                if(int(current_year) > latest_year and current_form_type == form_type): 
                    latest_year = current_year
                    data = current_tag_data
        except KeyError as e:
            pass
    return data

def get_filing_ids(json_data, form_type, num_filings):
    tags = constants.XBRL_tags["Net Income"]
    units = constants.Units_map["Net Income"]
    data = get_latest_data_from_tags(json_data, form_type, tags, units)

    filing_IDs = []
    for i in range(len(data)-1, -1, -1):
        if(len(filing_IDs) == num_filings): 
            break
        if("frame" in data[i] and len(data[i]["frame"]) > 6): 
            continue
        if(data[i]["form"] == form_type): 
            accn_number = data[i]["accn"]
            if(accn_number not in filing_IDs): 
                filing_IDs.append(accn_number)
    return filing_IDs

def get_quantity(tags, units, json_data, form_type, num_filings, filing_ids): 
    data = get_latest_data_from_tags(json_data, form_type, tags, units)   
    
    frame_length = 6
    if(form_type == "10-Q"): frame_length = 9

    values = []
    for i in range(len(data)-1, -1, -1):
        if(len(values) == num_filings): 
            break
        if("frame" in data[i] and len(data[i]["frame"]) != frame_length): 
            continue
        if(data[i]["form"] == form_type and data[i]["accn"] in filing_ids): 
            values.append(data[i]["val"])
    
    if(values == []): values = [0]*num_filings
    return values

def validate_inputs(inputs, num_filings):
    for i in range(len(inputs["Revenue"])): 
        if(inputs["Revenue"][i] == 0 or inputs["Operating Income"][i] == 0 or inputs["Net Income"][i] == 0):
            return constants.ERRORCODE
    if(inputs["COGS"] == [0]*num_filings):
        for i in range(num_filings):  
            cogs = (int(inputs["Revenue"][i]) - int(inputs["Operating Income"][i]))
            inputs["COGS"][i] = cogs
    if(inputs["Debt"] == [0]*num_filings):
        for i in range(num_filings): 
            debt = int(inputs["Current Long Term Debt"][i]) + int(inputs["Noncurrent Long Term Debt"][i])
            inputs["Debt"][i] = debt
    return inputs

def parse_inputs(json_data, form_type, num_filings):
    inputs = {}
    filing_ids = get_filing_ids(json_data, form_type, num_filings)
    assert(len(filing_ids) > 0), "Unable to find any filing IDs"

    for key in constants.XBRL_tags:
        units = constants.Units_map[key]
        tags = constants.XBRL_tags[key]
        inputs[key] = get_quantity(tags, units, json_data, form_type, num_filings, filing_ids)
    inputs = validate_inputs(inputs, num_filings)
    assert(inputs != constants.ERRORCODE), "One or more of revenue, operating Income or net Income were not loaded correctly"
    return inputs

def get_inputs(ticker, form_type="10-K", num_filings=1): 
    json_data = util.get_company_facts_json_from_ticker(ticker) 
    inputs = parse_inputs(json_data, form_type, num_filings)
    inputs["Price"] = [util.get_stock_price_from_ticker(ticker)]
    return inputs

if __name__ == "__main__":
    inputs = get_inputs("msft")
    print(inputs)