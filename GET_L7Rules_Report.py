import meraki, openpyxl, os

# Replace these with your Meraki API key and organization ID
api_key = os.getenv('API_Key_SDMeraki')

#Define orgs to skip
orgs_to_skip = ['Osborn Barr', 'Anders Minkler Huber & Helm', 'St. Louis Nephrology & Hypertension', 'Louisa Foods', 'The 816 Condominium', 'Yoon Dermatology']

# Create a Meraki API session
dashboard = meraki.DashboardAPI(api_key, suppress_logging=True)

# Initialize an Excel workbook
workbook = openpyxl.Workbook()
sheet = workbook.active
sheet.title = 'Current L7 Rules'

# Create headers in the Excel sheet
sheet.append(['Organization Name', 'Organization ID', 'Network ID', 'L7 Firewall Rule'])

# Step 1: Get organizations and their IDs
orgs = dashboard.organizations.getOrganizations()
for org in orgs:
    org_name = org['name']
    org_id = org['id']

    #Skip orgs if defined in orgs_to_skip
    if org_name in orgs_to_skip:
        continue

    # Step 2: Get networks owned by each organization
    networks = dashboard.organizations.getOrganizationNetworks(org_id)
    for network in networks:
        network_id = network['id']

        #check if it is an MX network; can only query MX appliances and not switches, etc
        if 'appliance' in network['productTypes']:

            # Step 3: Get L7 firewall rules for each network
            firewall_rules_data = dashboard.appliance.getNetworkApplianceFirewallL7FirewallRules(network_id)

            #Extract rules from the response
            firewall_rules = firewall_rules_data.get('rules', [])
        
            #iterate through the rules
            for rule in firewall_rules:
                policy = rule.get('policy', '')
                rule_type = rule.get('type', '')
                value = ', '.join(rule.get('value', []))
        
                # Write data to the Excel sheet
                sheet.append([org_name, org_id, network_id, policy, rule_type, value])

        else: #not an MX network
            sheet.append([org_name, org_id, network_id, 'Not an MX Network', '', ''])

# Save the Excel workbook
workbook.save('merakiReport_L7rules.xlsx')
