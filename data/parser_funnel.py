def parse_funnel(data, lc):
    rows = []
    
    # Safely navigate to the contents list
    contents = data.get("response", {}).get("data", {}).get("contents", [])
    
    for item in contents:
        funnel = item.get("funnel", {})
        steps = funnel.get("steps", [])
        
        # 1. Handle Steps logic for Website Match
        step_names = [step.get("event") for step in steps]
        tofu = step_names[0] if step_names else "N/A"
        bofu = step_names[-1] if len(step_names) > 1 else "N/A"
        
        # 2. Format "Steps" summary (e.g., "Step 1...Step 4")
        if len(step_names) > 1:
            steps_display = f"{tofu}...{bofu}"
        else:
            steps_display = tofu

        # 3. Conversion Window (e.g., "30 days")
        comp_time = funnel.get('completionTime', {})
        conversion_window = f"{comp_time.get('value', '')} {comp_time.get('type', '')}".strip()

        # 4. Construct Row to match Website Columns
        # Order: Report Name, Steps Count, TOFU, BOFU, Window, Modified On, Modified By, Created By
        rows.append([
            lc,                                     # Account
            funnel.get("title"),                    # Report (Website Match)
            len(steps),                             # Steps Count (Website Match)
            steps_display,                          # Steps Summary (Website Match)
            tofu,                                   # TOFU (Extra Detail)
            bofu,                                   # BOFU (Extra Detail)
            conversion_window,                      # Conversion Window (Website Match)
            "Last 30 days",                         # Reporting Window (Usually fixed in UI)
            funnel.get("lastModifiedAt"),           # Last Modified On (Website Match)
            funnel.get("lastModifiedBy"),           # Last Modified By (Website Match)
            funnel.get("createdBy"),                # Created By (Website Match)
            funnel.get("createdAt"),                # Created At
            funnel.get("id"),                       # Funnel ID
            funnel.get("status")                    # Status (ACTIVE/INACTIVE)
        ])

    return rows