import pandas as pd

def assign_projects(people_file, projects_file, output_file="final_assignments.xlsx"):
    # --- Load data ---
    people_df = pd.read_excel(people_file, sheet_name="People")
    projects_df = pd.read_excel(projects_file, sheet_name="Projects")
    
    # === Priority handling ===
    if "Priority" not in projects_df.columns:
        print("No 'Priority' column → assigning in file order")
        projects_df["Priority"] = 999
    
    projects_df = projects_df.sort_values(by="Priority", ascending=True).reset_index(drop=True)
    
    people_df["EffectiveCapacity"] = people_df["CapacityDays"] - people_df["VacationDays"]
    remaining_capacity = people_df.set_index("Name")["EffectiveCapacity"].to_dict()
    
    temp_assignments = []
    role_columns = [col for col in projects_df.columns if col not in ["Project", "Priority"]]
    
    print("Assigning projects in this order (lowest number = highest priority):")
    for i, proj in projects_df.iterrows():
        print(f" {proj['Priority']:>2} → {proj['Project']}")
    
    # --- Your original assignment loop (unchanged) ---
    for _, project_row in projects_df.iterrows():
        project_name = project_row["Project"]
        for role in role_columns:
            needed_days = project_row[role]
            if pd.isna(needed_days) or needed_days <= 0:
                continue
            needed_days = int(needed_days)
            
            candidates = people_df[
                people_df["Roles"].astype(str).str.contains(role, case=False, na=False)
            ].copy()
            
            if candidates.empty:
                print(f"Warning: No one has role '{role}' → {needed_days} days unassigned for '{project_name}' (Priority {project_row['Priority']})")
                continue
            
            while needed_days > 0:
                candidates["Remaining"] = candidates["Name"].map(remaining_capacity)
                candidates = candidates[candidates["Remaining"] > 0]
                if candidates.empty:
                    print(f"Warning: Not enough capacity left for '{role}' on '{project_name}' – {needed_days} days still needed")
                    break
                
                candidates = candidates.sort_values("Remaining")
                
                for _, person in candidates.iterrows():
                    if needed_days <= 0:
                        break
                    name = person["Name"]
                    if remaining_capacity[name] > 0:
                        temp_assignments.append({
                            "Project": project_name,
                            "Person": name,
                            "Role": role,
                            "Days": 1
                        })
                        remaining_capacity[name] -= 1
                        needed_days -= 1
    
    # === Build outputs ===
    if temp_assignments:
        temp_df = pd.DataFrame(temp_assignments)
        assignments_df = (temp_df
                          .groupby(["Project", "Person", "Role"], as_index=False)["Days"]
                          .sum()
                          .sort_values(["Project", "Role", "Person"])
                          .reset_index(drop=True))
        assigned_summary = temp_df.groupby(["Project", "Role"])["Days"].sum().reset_index(name="AssignedDays")
    else:
        assignments_df = pd.DataFrame(columns=["Project", "Person", "Role", "Days"])
        assigned_summary = pd.DataFrame(columns=["Project", "Role", "AssignedDays"])
    
    # Utilization Summary (per person)
    utilization = people_df[["Name", "EffectiveCapacity"]].copy()
    utilization["DaysUsed"] = 0
    if temp_assignments:
        days_per_person = temp_df.groupby("Person")["Days"].sum()
        utilization["DaysUsed"] = utilization["Name"].map(days_per_person).fillna(0).astype(int)
    utilization["Utilization %"] = (utilization["DaysUsed"] / utilization["EffectiveCapacity"] * 100).round(1)
    utilization = utilization[["Name", "EffectiveCapacity", "DaysUsed", "Utilization %"]]
    
    # ==================== NEW & IMPROVED: Utilization by Role ====================
    # 1. Total available capacity per role
    role_capacity = {}
    for role in role_columns:
        capable = people_df[people_df["Roles"].astype(str).str.contains(role, case=False, na=False)]
        role_capacity[role] = {
            "People": len(capable),
            "AvailableDays": int(capable["EffectiveCapacity"].sum())
        }
    
    # 2. Total required days per role (from Projects sheet)
    required_per_role = projects_df[role_columns].sum().to_dict()
    for role in required_per_role:
        required_per_role[role] = int(required_per_role[role]) if pd.notna(required_per_role[role]) else 0
    
    # 3. Actually assigned per role
    assigned_per_role = temp_df.groupby("Role")["Days"].sum().to_dict() if temp_assignments else {}
    
    # Build final table
    utilization_by_role = []
    for role in role_columns:
        available = role_capacity[role]["AvailableDays"]
        required = required_per_role.get(role, 0)
        assigned = assigned_per_role.get(role, 0)
        shortfall = assigned - available   # negative = we're short!
        remaining_capacity_role = available - assigned
        
        utilization_by_role.append({
            "Role": role,
            "People Capable": role_capacity[role]["People"],
            "Total Available Days": available,
            "Total Required Days": required,
            "Days Assigned": assigned,
            "Remaining Capacity": remaining_capacity_role,
            "Shortfall Days": shortfall,                    # ← This goes negative → bottleneck!
            "Utilization %": round(assigned / available * 100, 1) if available > 0 else 0,
            "Demand vs Supply %": round(required / available * 100, 1) if available > 0 else 999
        })
    
    utilization_by_role_df = pd.DataFrame(utilization_by_role)
    utilization_by_role_df = utilization_by_role_df.sort_values("Shortfall Days", ascending=True)  # worst first
    
    # Project Gaps & Summary (unchanged)
    required_long = projects_df.melt(
        id_vars=["Project", "Priority"],
        value_vars=role_columns,
        var_name="Role",
        value_name="RequiredDays"
    )
    required_long = required_long[required_long["RequiredDays"] > 0].copy()
    required_long["RequiredDays"] = required_long["RequiredDays"].astype(int)
    
    project_summary = required_long.merge(assigned_summary, on=["Project", "Role"], how="left")
    project_summary["AssignedDays"] = project_summary["AssignedDays"].fillna(0).astype(int)
    project_summary["MissingDays"] = project_summary["RequiredDays"] - project_summary["AssignedDays"]
    project_summary["Coverage %"] = (project_summary["AssignedDays"] / project_summary["RequiredDays"] * 100).round(1)
    project_summary = project_summary.sort_values(["Priority", "Project", "Role"])
    
    totals = project_summary.groupby(["Project", "Priority"]).agg({
        "RequiredDays": "sum",
        "AssignedDays": "sum",
        "MissingDays": "sum"
    }).reset_index()
    totals["Role"] = "TOTAL"
    totals["Coverage %"] = (totals["AssignedDays"] / totals["RequiredDays"] * 100).round(1)
    project_summary = pd.concat([project_summary, totals], ignore_index=True)
    
    # === Save everything ===
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        assignments_df.to_excel(writer, sheet_name="Assignments", index=False)
        utilization.to_excel(writer, sheet_name="Utilization Summary", index=False)
        project_summary.to_excel(writer, sheet_name="Project Gaps & Summary", index=False)
        utilization_by_role_df.to_excel(writer, sheet_name="Utilization by Role", index=False)
    
    total_missing = project_summary[project_summary["Role"] != "TOTAL"]["MissingDays"].sum()
    print(f"\nDONE! → {output_file}")
    print(f" • {len(assignments_df)} assignment lines")
    print(f" • {utilization['DaysUsed'].sum()} days assigned")
    print(f" • {total_missing} days still missing")
    print(f" • 'Utilization by Role' tab now shows negative Shortfall Days → instant bottleneck detection!")

    return assignments_df


if __name__ == "__main__":
    assign_projects("people_projects.xlsx", "people_projects.xlsx", "final_assignments.xlsx")