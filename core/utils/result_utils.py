"""
Utilities for processing rule results.
"""


def merge_results(results, merge_config):
    """
    Merge multiple rule results based on configuration.
    In-place merge - maintains original order of results.
    
    Args:
        results: List[dict] - List of rule execution results
            Each result: {
                "name": str,
                "mark": int/float,
                "errors": List[str],
                "needs_review": bool
            }
        merge_config: List[dict] - Merge configuration
            Each config: {
                "display_name": str,
                "rules": List[str],
                "show_details": bool (optional)
            }
    
    Returns:
        List[dict] - Merged results list (in original order)
    """
    if not merge_config:
        return results
    
    # Build lookup: rule_name -> (merge_def, position_in_merge)
    rule_to_merge = {}
    for merge_def in merge_config:
        rule_names = merge_def.get("rules", [])
        for idx, rule_name in enumerate(rule_names):
            rule_to_merge[rule_name] = (merge_def, idx)
    
    # Build results dict for quick lookup
    results_dict = {r["name"]: r for r in results}
    
    # Track which merge groups we've already processed
    processed_merges = set()
    
    # Build final result list in original order
    merged_results = []
    
    for result in results:
        rule_name = result["name"]
        
        # Check if this rule should be merged
        if rule_name in rule_to_merge:
            merge_def, position = rule_to_merge[rule_name]
            merge_key = merge_def.get("display_name", "Merged Rules")
            
            # Only process each merge group once (at first occurrence)
            if merge_key not in processed_merges:
                processed_merges.add(merge_key)
                
                # Collect all results for this merge group
                rule_names = merge_def.get("rules", [])
                show_details = merge_def.get("show_details", False)
                
                matching_results = [results_dict[name] for name in rule_names if name in results_dict]
                
                if matching_results:
                    # Calculate merged marks
                    total_mark = sum(r["mark"] for r in matching_results)
                    
                    # Collect all errors
                    all_errors = []
                    needs_review = False
                    for r in matching_results:
                        all_errors.extend(r.get("errors", []))
                        if r.get("needs_review", False):
                            needs_review = True
                    
                    # Create merged result
                    merged_result = {
                        "name": merge_key,
                        "mark": total_mark,
                        "errors": all_errors,
                        "needs_review": needs_review,
                        "merged": True,
                        "sub_rules": matching_results if show_details else []
                    }
                    
                    merged_results.append(merged_result)
            # else: skip this result (already merged at first occurrence)
        else:
            # Not part of any merge, keep as-is
            merged_results.append(result)
    
    return merged_results
