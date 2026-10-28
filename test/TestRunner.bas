Attribute VB_Name = "TestRunner"
' ASF test suite, standalone (no Rubberduck).
'   RunAllTests                          every test
'   RunTests "regex"                     one category
'   RunTests , "recursion_fib_single"    one test
' Failures and the summary go to the Immediate window. Every result goes to
' TestRunner.log next to the workbook; each test is logged before it starts,
' so if Excel dies the last line names the test that took it down.
Option Explicit
Private expected As Variant
Private actual As Variant
Private scriptEngine As ASF

Private pFilterCat As String, pFilterName As String
Private pCategory As String, pName As String
Private pAsserted As Boolean, pFailed As Boolean, pFailMsg As String
Private pRun As Long, pPassed As Long
Private pFailures As Collection
Private pLogPath As String
Private pStart As Single

Public Sub RunAllTests()
    RunTests
End Sub

Public Sub RunTests(Optional ByVal category As String, Optional ByVal testName As String)
    pFilterCat = LCase$(category): pFilterName = LCase$(testName)
    pRun = 0: pPassed = 0
    Set pFailures = New Collection
    OpenLog
    Dim t0 As Single: t0 = Timer
    WriteLog "=== ASF test suite " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & " ==="
    Debug.Print "=== ASF test suite ==="

    On Error GoTo Escaped
    If StartTest("arith",                     "arith_simple") Then arith_simple: EndTest
    If StartTest("arith",                     "arith_precedence") Then arith_precedence: EndTest
    If StartTest("paren",                     "paren_grouping") Then paren_grouping: EndTest
    If StartTest("negation",                  "negation_unary") Then negation_unary: EndTest
    If StartTest("power",                     "power_right_assoc") Then power_right_assoc: EndTest
    If StartTest("shortc",                    "shortc_and") Then shortc_and: EndTest
    If StartTest("shortc",                    "shortc_or") Then shortc_or: EndTest
    If StartTest("ternary",                   "ternary_operator") Then ternary_operator: EndTest
    If StartTest("shift",                     "left_shift") Then left_shift: EndTest
    If StartTest("shift",                     "right_shift") Then right_shift: EndTest
    If StartTest("compound_assignment",       "compound_assignment_plus_equals") Then compound_assignment_plus_equals: EndTest
    If StartTest("if",                        "if_chain_same_line") Then if_chain_same_line: EndTest
    If StartTest("if",                        "if_multiline") Then if_multiline: EndTest
    If StartTest("for",                       "for_simple") Then for_simple: EndTest
    If StartTest("for",                       "for_break_continue") Then for_break_continue: EndTest
    If StartTest("while",                     "while_break_continue") Then while_break_continue: EndTest
    If StartTest("switch",                    "switch_case") Then switch_case: EndTest
    If StartTest("try",                       "try_catch") Then try_catch: EndTest
    If StartTest("function",                  "function_basic") Then function_basic: EndTest
    If StartTest("function",                  "function_scope_isolation") Then function_scope_isolation: EndTest
    If StartTest("recursion",                 "recursion_fib_arrays") Then recursion_fib_arrays: EndTest
    If StartTest("recursion",                 "recursion_fib_single") Then recursion_fib_single: EndTest
    If StartTest("closure",                   "closure_shared_write") Then closure_shared_write: EndTest
    If StartTest("closure",                   "closure_multiple_instances") Then closure_multiple_instances: EndTest
    If StartTest("array",                     "array_literal_and_length") Then array_literal_and_length: EndTest
    If StartTest("array",                     "array_of_arrays_length") Then array_of_arrays_length: EndTest
    If StartTest("object_literal",            "object_literal_and_member") Then object_literal_and_member: EndTest
    If StartTest("nested_member",             "nested_member_index_LValue") Then nested_member_index_LValue: EndTest
    If StartTest("method_call",               "method_call_on_member") Then method_call_on_member: EndTest
    If StartTest("anon_func",                 "anon_func_as_arg") Then anon_func_as_arg: EndTest
    If StartTest("anon_func",                 "anon_func_closure_arg") Then anon_func_closure_arg: EndTest
    If StartTest("vbexpr",                    "mutating_VBAexpressions_Arrays") Then mutating_VBAexpressions_Arrays: EndTest
    If StartTest("vbexpr",                    "vbexpr_embedded") Then vbexpr_embedded: EndTest
    If StartTest("vbexpr",                    "reusing_results_from_VBAexpressions") Then reusing_results_from_VBAexpressions: EndTest
    If StartTest("calling_native",            "calling_native_function") Then calling_native_function: EndTest
    If StartTest("host_VBA_Expressions",      "host_sees_array_mutation") Then host_sees_array_mutation: EndTest
    If StartTest("map",                       "map_nested_array") Then map_nested_array: EndTest
    If StartTest("map",                       "map_array_of_objects") Then map_array_of_objects: EndTest
    If StartTest("map",                       "map_closure_capture") Then map_closure_capture: EndTest
    If StartTest("map",                       "map_returning_nested_array") Then map_returning_nested_array: EndTest
    If StartTest("map",                       "map_returning_objects_and_arrays") Then map_returning_objects_and_arrays: EndTest
    If StartTest("mapping",                   "mapping_mixed_types") Then mapping_mixed_types: EndTest
    If StartTest("filter",                    "filter_simple") Then filter_simple: EndTest
    If StartTest("filter",                    "filter_nested_arrays") Then filter_nested_arrays: EndTest
    If StartTest("reduce",                    "reduce_with_initial") Then reduce_with_initial: EndTest
    If StartTest("reduce",                    "reduce_with_NO_initial") Then reduce_with_NO_initial: EndTest
    If StartTest("slice",                     "slice_tail_only") Then slice_tail_only: EndTest
    If StartTest("slice",                     "slice_start_end") Then slice_start_end: EndTest
    If StartTest("pop_push",                  "pop_push") Then pop_push: EndTest
    If StartTest("pop_push",                  "push_nested_member") Then push_nested_member: EndTest
    If StartTest("range",                     "range_default") Then range_default: EndTest
    If StartTest("range",                     "range_custom") Then range_custom: EndTest
    If StartTest("range",                     "range_with_step") Then range_with_step: EndTest
    If StartTest("flatten",                   "flatten_full") Then flatten_full: EndTest
    If StartTest("flatten",                   "flatten_depth_one") Then flatten_depth_one: EndTest
    If StartTest("clone",                     "clone_array") Then clone_array: EndTest
    If StartTest("filter",                    "filter_reduce_chain") Then filter_reduce_chain: EndTest
    If StartTest("foreach",                   "foreach_method_updates_outer") Then foreach_method_updates_outer: EndTest
    If StartTest("variables",                 "variables_injection") Then variables_injection: EndTest
    If StartTest("variables",                 "variables_injection_collection") Then variables_injection_collection: EndTest
    If StartTest("foreach",                   "foreach_passes_index_and_array") Then foreach_passes_index_and_array: EndTest
    If StartTest("foreach",                   "foreach_builtin_signature_arrays") Then foreach_builtin_signature_arrays: EndTest
    If StartTest("foreach",                   "foreach_builtin_signature_objects") Then foreach_builtin_signature_objects: EndTest
    If StartTest("array",                     "array_unique_basic") Then array_unique_basic: EndTest
    If StartTest("array",                     "array_unique_nested") Then array_unique_nested: EndTest
    If StartTest("array",                     "array_concat") Then array_concat: EndTest
    If StartTest("array",                     "array_join_toString") Then array_join_toString: EndTest
    If StartTest("array",                     "array_shift_unshift") Then array_shift_unshift: EndTest
    If StartTest("array",                     "array_delete") Then array_delete: EndTest
    If StartTest("array",                     "array_splice_mutating") Then array_splice_mutating: EndTest
    If StartTest("array",                     "array_toSpliced_non_mutating") Then array_toSpliced_non_mutating: EndTest
    If StartTest("array",                     "array_at_negative") Then array_at_negative: EndTest
    If StartTest("array",                     "array_copyWithin") Then array_copyWithin: EndTest
    If StartTest("array",                     "array_entries") Then array_entries: EndTest
    If StartTest("array",                     "array_every") Then array_every: EndTest
    If StartTest("array",                     "array_find_and_indexes") Then array_find_and_indexes: EndTest
    If StartTest("array",                     "array_from_string") Then array_from_string: EndTest
    If StartTest("from",                      "from_array_copy") Then from_array_copy: EndTest
    If StartTest("from",                      "from_single_value_wrap") Then from_single_value_wrap: EndTest
    If StartTest("from",                      "from_with_map_array") Then from_with_map_array: EndTest
    If StartTest("from",                      "from_with_map_string") Then from_with_map_string: EndTest
    If StartTest("from",                      "from_nonclosure_second_arg_ignored") Then from_nonclosure_second_arg_ignored: EndTest
    If StartTest("array",                     "array_includes_indexOf_lastIndexOf") Then array_includes_indexOf_lastIndexOf: EndTest
    If StartTest("array",                     "array_of_factory_and_access") Then array_of_factory_and_access: EndTest
    If StartTest("array",                     "array_reverse_and_toReversed") Then array_reverse_and_toReversed: EndTest
    If StartTest("array",                     "array_some") Then array_some: EndTest
    If StartTest("array",                     "array_with_non_mutating") Then array_with_non_mutating: EndTest
    If StartTest("array",                     "array_sort_and_toSorted") Then array_sort_and_toSorted: EndTest
    If StartTest("array",                     "array_toSpliced_and_join") Then array_toSpliced_and_join: EndTest
    If StartTest("array",                     "array_entries_and_every_find_combo") Then array_entries_and_every_find_combo: EndTest
    If StartTest("array",                     "array_includes_and_index_checks_with_objects") Then array_includes_and_index_checks_with_objects: EndTest
    If StartTest("array",                     "array_complex_splice_and_copyWithin") Then array_complex_splice_and_copyWithin: EndTest
    If StartTest("deeply_nested",             "deeply_nested_Arrays") Then deeply_nested_Arrays: EndTest
    If StartTest("typeof",                    "typeof_array") Then typeof_array: EndTest
    If StartTest("typeof",                    "typeof_object") Then typeof_object: EndTest
    If StartTest("typeof",                    "typeof_closures") Then typeof_closures: EndTest
    If StartTest("for_in_of",                 "for_of_with_array") Then for_of_with_array: EndTest
    If StartTest("for_in_of",                 "for_in_with_array_indices") Then for_in_with_array_indices: EndTest
    If StartTest("for_in_of",                 "for_of_with_string") Then for_of_with_string: EndTest
    If StartTest("for_in_of",                 "for_in_with_object_properties") Then for_in_with_object_properties: EndTest
    If StartTest("string",                    "string_at") Then string_at: EndTest
    If StartTest("string",                    "string_at_negative_index") Then string_at_negative_index: EndTest
    If StartTest("string",                    "string_at_default_index") Then string_at_default_index: EndTest
    If StartTest("string",                    "string_length") Then string_length: EndTest
    If StartTest("string",                    "string_charAt") Then string_charAt: EndTest
    If StartTest("string",                    "string_charCodeAt") Then string_charCodeAt: EndTest
    If StartTest("string",                    "string_concat") Then string_concat: EndTest
    If StartTest("string",                    "string_endsWith") Then string_endsWith: EndTest
    If StartTest("string",                    "string_fromCharCode") Then string_fromCharCode: EndTest
    If StartTest("string",                    "string_includes") Then string_includes: EndTest
    If StartTest("string",                    "string_indexOf") Then string_indexOf: EndTest
    If StartTest("string",                    "string_lastIndexOf") Then string_lastIndexOf: EndTest
    If StartTest("string",                    "string_localCompare") Then string_localCompare: EndTest
    If StartTest("string",                    "string_pad") Then string_pad: EndTest
    If StartTest("string",                    "string_template") Then string_template: EndTest
    If StartTest("string",                    "string_template_with_arrays") Then string_template_with_arrays: EndTest
    If StartTest("string",                    "string_replace") Then string_replace: EndTest
    If StartTest("string",                    "string_slice") Then string_slice: EndTest
    If StartTest("string",                    "string_splite") Then string_splite: EndTest
    If StartTest("string",                    "string_startWith") Then string_startWith: EndTest
    If StartTest("string",                    "string_subString") Then string_subString: EndTest
    If StartTest("string",                    "string_toLowercase") Then string_toLowercase: EndTest
    If StartTest("string",                    "string_toUppercase") Then string_toUppercase: EndTest
    If StartTest("string",                    "string_trim") Then string_trim: EndTest
    If StartTest("string",                    "string_trim_start_end") Then string_trim_start_end: EndTest
    If StartTest("regex",                     "regex_replace_from_string_object") Then regex_replace_from_string_object: EndTest
    If StartTest("regex",                     "regex_replacer_function_from_string_object") Then regex_replacer_function_from_string_object: EndTest
    If StartTest("regex",                     "regex_replace_using_placeholders_from_string_object") Then regex_replace_using_placeholders_from_string_object: EndTest
    If StartTest("regex",                     "regex_editing_matches_from_string_object") Then regex_editing_matches_from_string_object: EndTest
    If StartTest("regex",                     "regex_replace_using_templates_from_string_object") Then regex_replace_using_templates_from_string_object: EndTest
    If StartTest("regex",                     "regex_match_from_string_object") Then regex_match_from_string_object: EndTest
    If StartTest("regex",                     "regex_match_global_from_string_object") Then regex_match_global_from_string_object: EndTest
    If StartTest("regex",                     "regex_matchall_from_string_object") Then regex_matchall_from_string_object: EndTest
    If StartTest("regex",                     "regex_replace_empty_alternative_global") Then regex_replace_empty_alternative_global: EndTest
    If StartTest("regex",                     "regex_replace_empty_match_keeps_chars") Then regex_replace_empty_match_keeps_chars: EndTest
    If StartTest("regex",                     "regex_replace_lookahead_insertion") Then regex_replace_lookahead_insertion: EndTest
    If StartTest("regex",                     "regex_replace_empty_match_at_start") Then regex_replace_empty_match_at_start: EndTest
    If StartTest("regex",                     "regex_match_global_empty_alternative") Then regex_match_global_empty_alternative: EndTest
    If StartTest("regex",                     "regex_replaceall_requires_global_flag") Then regex_replaceall_requires_global_flag: EndTest
    If StartTest("regex",                     "regex_replaceall_with_global_flag") Then regex_replaceall_with_global_flag: EndTest
    If StartTest("regex",                     "regex_object_constructor") Then regex_object_constructor: EndTest
    If StartTest("regex",                     "regex_object_execute_method") Then regex_object_execute_method: EndTest
    If StartTest("regex",                     "regex_object_replace_method") Then regex_object_replace_method: EndTest
    If StartTest("regex",                     "regex_object_split_method") Then regex_object_split_method: EndTest
    If StartTest("regex",                     "regex_object_executeAll_method") Then regex_object_executeAll_method: EndTest
    If StartTest("regex",                     "regex_object_method_with_flag_property") Then regex_object_method_with_flag_property: EndTest
    If StartTest("regex",                     "regex_object_named_captures") Then regex_object_named_captures: EndTest
    If StartTest("regex",                     "regex_object_conditional_simple") Then regex_object_conditional_simple: EndTest
    If StartTest("regex",                     "regex_object_conditional_inline_flags") Then regex_object_conditional_inline_flags: EndTest
    If StartTest("regex",                     "regex_object_conditional_inline_flags_off") Then regex_object_conditional_inline_flags_off: EndTest
    If StartTest("regex",                     "regex_object_inline_dotAll_flags") Then regex_object_inline_dotAll_flags: EndTest
    If StartTest("regex",                     "regex_object_inline_dotAll_flags_off") Then regex_object_inline_dotAll_flags_off: EndTest
    If StartTest("regex",                     "regex_object_Named_Conditionals") Then regex_object_Named_Conditionals: EndTest
    If StartTest("regex",                     "regex_object_Named_Conditionals_NO_Match") Then regex_object_Named_Conditionals_NO_Match: EndTest
    If StartTest("classes",                   "classes_static_method_declarations") Then classes_static_method_declarations: EndTest
    If StartTest("classes",                   "classes_field_declarations") Then classes_field_declarations: EndTest
    If StartTest("classes",                   "classes_inheritance_with_fields") Then classes_inheritance_with_fields: EndTest
    If StartTest("classes",                   "classes_fields_declarations_with_mixed_syntax") Then classes_fields_declarations_with_mixed_syntax: EndTest
    If StartTest("classes",                   "classes_fields_with_complex_initializers") Then classes_fields_with_complex_initializers: EndTest
    If StartTest("classes",                   "classes_multiple_instances") Then classes_multiple_instances: EndTest
    If StartTest("classes",                   "classes_method_chaining") Then classes_method_chaining: EndTest
    If StartTest("object",                    "object_keys") Then object_keys: EndTest
    If StartTest("object",                    "object_values") Then object_values: EndTest
    If StartTest("object",                    "object_entries") Then object_entries: EndTest
    If StartTest("object",                    "object_size") Then object_size: EndTest
    If StartTest("object",                    "object_length") Then object_length: EndTest
    If StartTest("object",                    "object_hasKey_true") Then object_hasKey_true: EndTest
    If StartTest("object",                    "object_hasKey_false") Then object_hasKey_false: EndTest
    If StartTest("object",                    "object_has_alias") Then object_has_alias: EndTest
    If StartTest("object",                    "object_isEmpty_true") Then object_isEmpty_true: EndTest
    If StartTest("object",                    "object_isEmpty_false") Then object_isEmpty_false: EndTest
    If StartTest("object",                    "object_get_existing_key") Then object_get_existing_key: EndTest
    If StartTest("object",                    "object_get_with_default") Then object_get_with_default: EndTest
    If StartTest("object",                    "object_get_without_default_missing_key") Then object_get_without_default_missing_key: EndTest
    If StartTest("object",                    "object_set_new_key") Then object_set_new_key: EndTest
    If StartTest("object",                    "object_set_update_existing") Then object_set_update_existing: EndTest
    If StartTest("object",                    "object_delete_existing_key") Then object_delete_existing_key: EndTest
    If StartTest("object",                    "object_remove_alias") Then object_remove_alias: EndTest
    If StartTest("object",                    "object_clear") Then object_clear: EndTest
    If StartTest("object",                    "object_clone_simple") Then object_clone_simple: EndTest
    If StartTest("object",                    "object_clone_nested") Then object_clone_nested: EndTest
    If StartTest("object",                    "object_merge") Then object_merge: EndTest
    If StartTest("object",                    "object_merge_nested") Then object_merge_nested: EndTest
    If StartTest("object",                    "object_forEach") Then object_forEach: EndTest
    If StartTest("object",                    "object_forEach_with_key") Then object_forEach_with_key: EndTest
    If StartTest("object",                    "object_map") Then object_map: EndTest
    If StartTest("object",                    "object_map_with_key") Then object_map_with_key: EndTest
    If StartTest("object",                    "object_filter") Then object_filter: EndTest
    If StartTest("object",                    "object_filter_by_key") Then object_filter_by_key: EndTest
    If StartTest("object",                    "object_some_true") Then object_some_true: EndTest
    If StartTest("object",                    "object_some_false") Then object_some_false: EndTest
    If StartTest("object",                    "object_every_true") Then object_every_true: EndTest
    If StartTest("object",                    "object_every_false") Then object_every_false: EndTest
    If StartTest("object",                    "object_chaining_methods") Then object_chaining_methods: EndTest
    If StartTest("object",                    "object_with_nested_arrays") Then object_with_nested_arrays: EndTest
    If StartTest("object",                    "object_entries_with_forEach") Then object_entries_with_forEach: EndTest
    If StartTest("polymorphism",              "polymorphism_basic") Then polymorphism_basic: EndTest
    If StartTest("polymorphism",              "polymorphism_three_levels") Then polymorphism_three_levels: EndTest
    If StartTest("polymorphism",              "polymorphism_method_delegation") Then polymorphism_method_delegation: EndTest
    If StartTest("sort",                      "sort_chain_on_objects") Then sort_chain_on_objects: EndTest
    If StartTest("injecting",                 "injecting_asf_objects") Then injecting_asf_objects: EndTest
    If StartTest("spread",                    "spread_simple") Then spread_simple: EndTest
    If StartTest("spread",                    "spread_multiple") Then spread_multiple: EndTest
    If StartTest("spread",                    "spread_with_empty") Then spread_with_empty: EndTest
    If StartTest("spread",                    "spread_strings") Then spread_strings: EndTest
    If StartTest("spread",                    "spread_mixed_string") Then spread_mixed_string: EndTest
    If StartTest("spread",                    "spread_combined") Then spread_combined: EndTest
    If StartTest("spread",                    "spread_nested") Then spread_nested: EndTest
    If StartTest("spread",                    "spread_shallow_copy") Then spread_shallow_copy: EndTest
    If StartTest("spread",                    "spread_basic_function_call") Then spread_basic_function_call: EndTest
    If StartTest("spread",                    "spread_many_in_function_call") Then spread_many_in_function_call: EndTest
    If StartTest("spread",                    "spread_objects_basic") Then spread_objects_basic: EndTest
    If StartTest("spread",                    "spread_objects_override") Then spread_objects_override: EndTest
    If StartTest("spread",                    "spread_multiple_objects") Then spread_multiple_objects: EndTest
    If StartTest("spread",                    "spread_array_in_object") Then spread_array_in_object: EndTest
    If StartTest("spread",                    "spread_string_in_object") Then spread_string_in_object: EndTest
    If StartTest("spread",                    "spread_objects_with_primitives") Then spread_objects_with_primitives: EndTest
    If StartTest("rest_parameter",            "rest_parameter_basic") Then rest_parameter_basic: EndTest
    If StartTest("rest_parameter",            "rest_parameter_no_arguments") Then rest_parameter_no_arguments: EndTest
    If StartTest("rest_parameter",            "rest_parameter_with_regular_parameters") Then rest_parameter_with_regular_parameters: EndTest
    If StartTest("rest_parameter",            "rest_parameter_with_single_extra_argument") Then rest_parameter_with_single_extra_argument: EndTest
    If StartTest("rest_parameter",            "rest_parameter_not_at_end") Then rest_parameter_not_at_end: EndTest
    If StartTest("rest_parameter",            "rest_parameter_with_exact_number_of_arguments") Then rest_parameter_with_exact_number_of_arguments: EndTest
    If StartTest("rest_parameter",            "rest_parameter_receiving_expread_arguments") Then rest_parameter_receiving_expread_arguments: EndTest
    If StartTest("rest_parameter",            "rest_parameter_and_expread_chain") Then rest_parameter_and_expread_chain: EndTest
    If StartTest("array_destructuring",       "array_destructuring_basic") Then array_destructuring_basic: EndTest
    If StartTest("array_destructuring",       "array_destructuring_fewer_targets") Then array_destructuring_fewer_targets: EndTest
    If StartTest("array_destructuring",       "array_destructuring_more_targets") Then array_destructuring_more_targets: EndTest
    If StartTest("array_destructuring",       "array_destructuring_with_rest") Then array_destructuring_with_rest: EndTest
    If StartTest("array_destructuring",       "array_destructuring_with_no_remaining_elm") Then array_destructuring_with_no_remaining_elm: EndTest
    If StartTest("module_system",             "module_system_working_directory") Then module_system_working_directory: EndTest
    If StartTest("module_system",             "module_system_named_exports_imports") Then module_system_named_exports_imports: EndTest
    If StartTest("module_system",             "module_system_default_export") Then module_system_default_export: EndTest
    If StartTest("module_system",             "module_system_namespace_import") Then module_system_namespace_import: EndTest
    If StartTest("module_system",             "module_system_mixed_imports") Then module_system_mixed_imports: EndTest
    If StartTest("module_system",             "module_system_prototype_imports") Then module_system_prototype_imports: EndTest
    If StartTest("Modules",                   "module_system_prototype_single_import") Then module_system_prototype_single_import: EndTest
    If StartTest("Modules",                   "module_system_prototype_namespace_import") Then module_system_prototype_namespace_import: EndTest
    If StartTest("Modules",                   "module_system_prototype_side_effect_import") Then module_system_prototype_side_effect_import: EndTest
    If StartTest("math",                      "math_basic") Then math_basic: EndTest
    If StartTest("calling_native_app_level",  "calling_native_function_app_level") Then calling_native_function_app_level: EndTest
    If StartTest("calling_native_app_level",  "calling_native_function_app_level_2") Then calling_native_function_app_level_2: EndTest
    If StartTest("calling_native_app_level",  "calling_array_method_from_collection") Then calling_array_method_from_collection: EndTest
    If StartTest("advanced_prototyping",      "monkey_patch_worksheets_name") Then monkey_patch_worksheets_name: EndTest
    If StartTest("Nullish coalescing",        "nullish_coalescing") Then nullish_coalescing: EndTest
    On Error GoTo 0

    Dim summary As String
    summary = "=== " & pPassed & "/" & pRun & " passed, " & pFailures.Count & " failed (" & Format$(Timer - t0, "0.00") & " s) ==="
    WriteLog summary
    Debug.Print summary
    If pLogPath <> vbNullString Then Debug.Print "Log: " & pLogPath
    Set scriptEngine = Nothing
    Exit Sub
Escaped:
    ' an error the test did not trap: fail it and go on with EndTest
    AssertFail "Unhandled error #" & err.Number & " - " & err.Description
    Resume Next
End Sub

Private Function StartTest(ByVal category As String, ByVal testName As String) As Boolean
    If pFilterCat <> vbNullString And LCase$(category) <> pFilterCat Then Exit Function
    If pFilterName <> vbNullString And LCase$(testName) <> pFilterName Then Exit Function
    pCategory = category: pName = testName
    pAsserted = False: pFailed = False: pFailMsg = vbNullString
    expected = Empty: actual = Empty
    Set scriptEngine = Nothing
    WriteLog "RUN  " & category & "/" & testName
    DoEvents
    pStart = Timer
    StartTest = True
End Function

Private Sub EndTest()
    Dim ms As String: ms = " (" & Format$((Timer - pStart) * 1000, "0") & " ms)"
    If Not pAsserted Then AssertFail "No assertion was made"
    pRun = pRun + 1
    If pFailed Then
        Dim entry As String: entry = "FAIL " & pCategory & "/" & pName & ": " & pFailMsg
        pFailures.Add entry
        WriteLog entry & ms
        Debug.Print entry
    Else
        pPassed = pPassed + 1
        WriteLog "PASS " & pCategory & "/" & pName & ms
    End If
    Set scriptEngine = Nothing
End Sub

' Strict like Rubberduck's AssertClass: types must match, strings compare case-sensitively
Private Sub AssertEqual(ByVal expectedValue As Variant, ByVal actualValue As Variant)
    pAsserted = True
    If IsEmpty(expectedValue) Then expectedValue = vbNullString
    If IsEmpty(actualValue) Then actualValue = vbNullString
    If IsObject(expectedValue) Or IsObject(actualValue) Then
        If IsObject(expectedValue) And IsObject(actualValue) Then
            If expectedValue Is actualValue Then Exit Sub
        End If
    ElseIf VarType(expectedValue) <> VarType(actualValue) Then
        AssertFail "Type mismatch, expected " & Describe(expectedValue) & ", actual " & Describe(actualValue)
        Exit Sub
    ElseIf IsNull(expectedValue) Then
        Exit Sub
    ElseIf Not IsArray(expectedValue) Then
        If expectedValue = actualValue Then Exit Sub
    End If
    AssertFail "Expected " & Describe(expectedValue) & ", actual " & Describe(actualValue)
End Sub

Private Sub AssertFail(ByVal message As String)
    pAsserted = True
    If pFailed Then Exit Sub
    pFailed = True
    pFailMsg = message
End Sub

Private Function Describe(ByRef v As Variant) As String
    If IsObject(v) Then
        Describe = "<" & TypeName(v) & ">"
    ElseIf IsArray(v) Or IsNull(v) Then
        Describe = TypeName(v)
    Else
        Describe = TypeName(v) & " <" & ConvertNewLines(CStr(v)) & ">"
    End If
End Function

Private Sub OpenLog()
    pLogPath = vbNullString
    On Error GoTo NoLog
    If ThisWorkbook.path = vbNullString Then Exit Sub
    pLogPath = ThisWorkbook.path & Application.PathSeparator & "TestRunner.log"
    Dim f As Integer: f = FreeFile
    Open pLogPath For Output As #f
    Close #f
    Exit Sub
NoLog:
    pLogPath = vbNullString
End Sub

' Opened and closed per line so the file survives a crash
Private Sub WriteLog(ByVal msg As String)
    If pLogPath = vbNullString Then Exit Sub
    On Error GoTo NoLog
    Dim f As Integer: f = FreeFile
    Open pLogPath For Append As #f
    Print #f, msg
    Close #f
    Exit Sub
NoLog:
    pLogPath = vbNullString
End Sub

Private Function GetResult(script As String, Optional verbose As Boolean = False) As Variant
    On Error Resume Next
    Dim idx As Long
    Set scriptEngine = New ASF
    
    With scriptEngine
        scriptEngine.verbose = verbose
        idx = .Compile(script)
        .Run idx
        GetResult = .OUTPUT_
    End With
End Function
Private Function ConvertNewLines(aStr As String) As String
    ConvertNewLines = VBA.Replace(VBA.Replace(aStr, vbLf, "\n"), vbCr, "\r")
End Function
Private Sub arith_simple()
    On Error GoTo TestFail
    
    actual = CStr(GetResult("return(1 + 2 * 3);"))
    expected = "7"
    AssertEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub arith_precedence()
    On Error GoTo TestFail
    
    actual = CStr(GetResult("return(1 + 2 * 3 / 4^2);"))
    expected = "1.375"
    AssertEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub paren_grouping()
    On Error GoTo TestFail
    
    actual = CStr(GetResult("return((1 + 2) * 3);"))
    expected = "9"
    AssertEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub negation_unary()
    On Error GoTo TestFail
    Dim Globals As ASF_Globals
    GetResult "print(-5 + 3, !false, !true);", True
    Set Globals = scriptEngine.GetGlobals
    With Globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.Count))
    End With
    expected = "PRINT:-2, True, False"
    AssertEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub power_right_assoc()
    On Error GoTo TestFail
    
    actual = CStr(GetResult("return(2 ^ 3 ^ 2);"))
    expected = "512"
    AssertEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub shortc_and()
    On Error GoTo TestFail
    
    actual = CBool(GetResult("x = false; return(x && (1/0));"))
    expected = CBool("False")
    AssertEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub shortc_or()
    On Error GoTo TestFail
    
    actual = CBool(GetResult("x = true; return(x || (1/0));"))
    expected = CBool("True")
    AssertEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub ternary_operator()
    On Error GoTo TestFail
    
    actual = CStr(GetResult("return( 1 < 2 ? 'yes' : 'no' )"))
    expected = "yes"
    AssertEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub left_shift()
    On Error GoTo TestFail
    
    actual = CStr(GetResult("return(5<<1)"))
    expected = "10"
    AssertEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub right_shift()
    On Error GoTo TestFail
    
    actual = CStr(GetResult("x=-100; x>>=5; return(x)"))
    expected = "-4"
    AssertEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub compound_assignment_plus_equals()
    On Error GoTo TestFail
    
    actual = CStr(GetResult("a=2; a += 3; return(a);"))
    expected = "5"
    AssertEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub if_chain_same_line()
    On Error GoTo TestFail
    Dim Globals As ASF_Globals
    GetResult "a=2; if (a==1) { print('one') } elseif (a==2) { print('two') } else { print('other') }; print('done');", True
    Set Globals = scriptEngine.GetGlobals
    With Globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.Count - 1)) & ", " & CStr(.gRuntimeLog(.gRuntimeLog.Count))
    End With
    expected = "PRINT:'two', PRINT:'done'"
    AssertEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub if_multiline()
    On Error GoTo TestFail
    Dim Globals As ASF_Globals
    GetResult "a=3;" & _
                    "if (a==1) {" & _
                    "  print('one')" & _
                    "} elseif (a==2) {" & _
                    "  print('two')" & _
                    "} elseif (a==3) {" & _
                    "  print('three')" & _
                    "} else {" & _
                    "  print('other')" & _
                    "};" & _
                    "print('end');", True
    Set Globals = scriptEngine.GetGlobals
    With Globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.Count - 1)) & ", " & CStr(.gRuntimeLog(.gRuntimeLog.Count))
    End With
    expected = "PRINT:'three', PRINT:'end'"
    AssertEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub for_simple()
    On Error GoTo TestFail
    
    actual = CStr(GetResult("s=0; for(i=1,i<=3,i=i+1) { s = s + i }; return(s);"))
    expected = "6"
    AssertEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub for_break_continue()
    On Error GoTo TestFail
    
    actual = CStr(GetResult("s=0; for(i=1,i<=5,i=i+1) {" & _
                                    "if (i==3) { continue }" & _
                                    "if (i==5) { break } s = s + i };" & _
                                    "return(s);"))
    expected = "7"
    AssertEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub while_break_continue()
    On Error GoTo TestFail
    
    actual = CStr(GetResult("i=1; s=0; while (i <= 5) {" & _
                                        "if (i==2) { i = i + 1 ; continue }" & _
                                        "if (i==5) { break }" & _
                                        "s = s + i ; i = i + 1 };" & _
                                        "return(s);"))
    expected = "8"
    AssertEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub switch_case()
    On Error GoTo TestFail
    
    actual = CStr(GetResult("c='blue'; switch(c) {" & _
                                        "case 'red' { return('warm') }" & _
                                        "case 'blue' { return('cool') }" & _
                                        "default { return('other') } }"))
    expected = "cool"
    AssertEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub try_catch()
    On Error GoTo TestFail
    
    actual = CStr(GetResult("try { x = 1/0 }" & _
                            "catch { return('caught') }"))
    expected = "caught"
    AssertEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub function_basic()
    On Error GoTo TestFail
    
    actual = CStr(GetResult("fun add(a,b) { return a + b }; return(add(2,3));"))
    expected = "5"
    AssertEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub function_scope_isolation()
    On Error GoTo TestFail
    Dim Globals As ASF_Globals
    GetResult "a=5; fun f(a) { a = a + 1 ; print(a) } ; f(a); print(a);", True
    Set Globals = scriptEngine.GetGlobals
    With Globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.Count - 1)) & ", " & CStr(.gRuntimeLog(.gRuntimeLog.Count))
    End With
    expected = "PRINT:6, PRINT:5"
    AssertEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub recursion_fib_arrays()
    On Error GoTo TestFail
    Dim Globals As ASF_Globals
    GetResult "fun fib(n) {" & _
                            " if (n <= 2) { return 1 }; return fib(n-1) + fib(n-2)" & _
                            "} ;" & _
                            "a = [];" & _
                            "for(i=1,i<=6,i=i+1) {" & _
                                "a[i] = fib(i)" & _
                            "};" & _
                            "print(a[1]); print(a[6]);" & _
                            "print(a);", True
    Set Globals = scriptEngine.GetGlobals
    With Globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.Count - 2)) & ", " _
                    & CStr(.gRuntimeLog(.gRuntimeLog.Count - 1)) & ", " _
                    & CStr(.gRuntimeLog(.gRuntimeLog.Count))
    End With
    expected = "PRINT:1, PRINT:8, PRINT:[ 1, 1, 2, 3, 5, 8 ]"
    AssertEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub recursion_fib_single()
    On Error GoTo TestFail
    
    actual = CStr(GetResult("fun fib(n) {" & _
                                " if (n <= 2) { return 1 }; return fib(n-1) + fib(n-2)" & _
                            "} ;" & _
                            "return(fib(15));"))
    expected = "610"
    AssertEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub closure_shared_write()
    On Error GoTo TestFail
    Dim Globals As ASF_Globals
    GetResult "a = 1; f = fun() { a = a + 1; return a }; print(f()); print(a);", True
    Set Globals = scriptEngine.GetGlobals
    With Globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.Count - 1)) & ", " _
                    & CStr(.gRuntimeLog(.gRuntimeLog.Count))
    End With
    expected = "PRINT:2, PRINT:2"
    AssertEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub closure_multiple_instances()
    On Error GoTo TestFail
    Dim Globals As ASF_Globals
    GetResult "a = 0; fun make() { return fun() { a = a + 1 ; return a } };" & _
                "f1 = make(); f2 = make();" & _
                "print(f1()); print(f2()); print(a);", True
    Set Globals = scriptEngine.GetGlobals
    With Globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.Count - 2)) & ", " _
                    & CStr(.gRuntimeLog(.gRuntimeLog.Count - 1)) & ", " _
                    & CStr(.gRuntimeLog(.gRuntimeLog.Count))
    End With
    expected = "PRINT:1, PRINT:2, PRINT:2"
    AssertEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub array_literal_and_length()
    On Error GoTo TestFail
    Dim Globals As ASF_Globals
    GetResult "a=[10,20,30]; print(a[2]); print(a.length);", True
    Set Globals = scriptEngine.GetGlobals
    With Globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.Count - 1)) & ", " _
                    & CStr(.gRuntimeLog(.gRuntimeLog.Count))
    End With
    expected = "PRINT:20, PRINT:3"
    AssertEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub array_of_arrays_length()
    On Error GoTo TestFail
    Dim Globals As ASF_Globals
    GetResult "a = [] ; a[1] = [7,8] ; a[3] = [9,10,11] ;" & _
                "print(a[1]); print(a[3]); print(a[3].length)", True
    Set Globals = scriptEngine.GetGlobals
    With Globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.Count - 2)) & ", " _
                    & CStr(.gRuntimeLog(.gRuntimeLog.Count - 1)) & ", " _
                    & CStr(.gRuntimeLog(.gRuntimeLog.Count))
    End With
    expected = "PRINT:[ 7, 8 ], PRINT:[ 9, 10, 11 ], PRINT:3"
    AssertEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub object_literal_and_member()
    On Error GoTo TestFail
    Dim Globals As ASF_Globals
    GetResult "o = { x: 10, y: 'hi' } ; print(o.x) ; o.x = o.x + 5 ; print(o.x)", True
    Set Globals = scriptEngine.GetGlobals
    With Globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.Count - 1)) & ", " _
                    & CStr(.gRuntimeLog(.gRuntimeLog.Count))
    End With
    expected = "PRINT:10, PRINT:15"
    AssertEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub nested_member_index_LValue()
    On Error GoTo TestFail
    
    actual = CStr(GetResult(" o = { a: [ {v:1}, {v:2} ] } ;" & _
                "o.a[2].v = o.a[2].v + 5 ; return(o.a[2].v + 2)"))
    expected = "9"
    AssertEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub method_call_on_member()
    On Error GoTo TestFail
    
    actual = CStr(GetResult("o = { v: 10, incr: fun(x) { return x + 1 } } ; return(o.incr(o.v))"))
    expected = "11"
    AssertEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub anon_func_as_arg()
    On Error GoTo TestFail
    
    actual = CStr(GetResult("fun apply(f,x) { return f(x) } ; return(apply(fun(y) { return y * 2 }, 5))"))
    expected = "10"
    AssertEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub anon_func_closure_arg()
    On Error GoTo TestFail
    
    actual = CStr(GetResult("a = 5; fun apply(f) { return f() }; return(apply(fun() { return a + 1 }))"))
    expected = "6"
    AssertEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub mutating_VBAexpressions_Arrays()
    On Error GoTo TestFail
    Dim Globals As ASF_Globals
    GetResult "a=@({{1;2;3};{4;(5+4);'value'}}); a[1]=2*5; print(a);", True
    Set Globals = scriptEngine.GetGlobals
    With Globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.Count))
    End With
    expected = "PRINT:[ 10, [ 4, 9, 'value' ] ]"
    AssertEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub vbexpr_embedded()
    On Error GoTo TestFail
    Dim Globals As ASF_Globals
    GetResult "a = @({1;0;4});" & _
                            " b = @({1;1;6});" & _
                            " c = @({-3;0;-10});" & _
                            " print(@(MROUND(LUDECOMP(ARRAY(a;b;c));4)))", True
    Set Globals = scriptEngine.GetGlobals
    With Globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.Count))
    End With
    expected = "PRINT:[ [ -3, 0, -10 ], [ -0.3333, 1, 2.6667 ], [ -0.3333, 0, 0.6667 ] ]"
    AssertEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub reusing_results_from_VBAexpressions()
    On Error GoTo TestFail
    
    actual = CStr(GetResult("a=@(lgn(32;2)); b=a*2; return b;"))
    expected = "10"
    AssertEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub calling_native_function()
    On Error GoTo TestFail
    Dim asfGlobals As New ASF_Globals
    Dim progIdx  As Long
    
    With asfGlobals
        .ASF_InitGlobals
        .gExprEvaluator.DeclareUDF "ThisWBname", "UserDefFunctions"
    End With
    Set scriptEngine = New ASF
    With scriptEngine
        .SetGlobals asfGlobals
        progIdx = .Compile("/*Get Thisworkbook name*/ return(@(ThisWBname()))")
        .Run progIdx
        actual = CStr(.OUTPUT_)
    End With
    expected = ThisWorkbook.name
    AssertEqual expected, actual
    
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub host_sees_array_mutation()
    On Error GoTo TestFail
    Dim asfGlobals As New ASF_Globals
    Dim progIdx  As Long
    
    With asfGlobals
        .ASF_InitGlobals
        .gExprEvaluator.DeclareUDF "HostCheckSecondNestedElement", "UserDefFunctions"
    End With
    Set scriptEngine = New ASF
    With scriptEngine
        .SetGlobals asfGlobals
        progIdx = .Compile("a=@({{7;8;9}}); a[1][2]=42; return(@(HostCheckSecondNestedElement(a)));")
        .Run progIdx
        actual = CStr(.OUTPUT_)
    End With
    expected = "42"
    AssertEqual expected, actual
    
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub map_nested_array()
    On Error GoTo TestFail
    Dim Globals As ASF_Globals
    GetResult "a = [1,[2,[3,[4]]]]; b = a.map(fun(x) { return x * 10 }); print(b);", True
    Set Globals = scriptEngine.GetGlobals
    With Globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.Count))
    End With
    expected = "PRINT:[ 10, [ 20, [ 30, [ 40 ] ] ] ]"
    AssertEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub map_array_of_objects()
    On Error GoTo TestFail
    Dim Globals As ASF_Globals
    GetResult "a = [{ k: 1, arr: [10,20] }, { k: 2, arr: [30,[40,50]] }];" & _
                        "b = a.map(fun(o){return { k: o.k * 2, arr: o.arr.map(fun(x){ return x + 1 })}; });" & _
                        "print(b);", True
    Set Globals = scriptEngine.GetGlobals
    With Globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.Count))
    End With
    expected = "PRINT:[ { k: 2, arr: [ 11, 21 ] }, { k: 4, arr: [ 31, [ 41, 51 ] ] } ]"
    AssertEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub map_closure_capture()
    On Error GoTo TestFail
    Dim Globals As ASF_Globals
    GetResult "mul = fun(factor){return fun(x){ return x * factor };};" & _
                        "a = [1,2,3]; b = a.map(mul(5));" & _
                        "print(b);", True
    Set Globals = scriptEngine.GetGlobals
    With Globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.Count))
    End With
    expected = "PRINT:[ 5, 10, 15 ]"
    AssertEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub map_returning_nested_array()
    On Error GoTo TestFail
    Dim Globals As ASF_Globals
    GetResult "print( [1,2].map(fun(x){ return [x,x] }) );", True
    Set Globals = scriptEngine.GetGlobals
    With Globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.Count))
    End With
    expected = "PRINT:[ [ 1, 1 ], [ 2, 2 ] ]"
    AssertEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub map_returning_objects_and_arrays()
    On Error GoTo TestFail
    Dim Globals As ASF_Globals
    GetResult "a = [1,2];" & _
                        "b = a.map(fun(n){return {orig: n,pair: [n, n*n],nested: [ [n, n+1], { v: n*n } ]};});" & _
                        "print(b);", True
    Set Globals = scriptEngine.GetGlobals
    With Globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.Count))
    End With
    expected = "PRINT:[ { orig: 1, pair: [ 1, 1 ], nested: [ [ 1, 2 ], { v: 1 } ] }, { orig: 2, pair: [ 2, 4 ], nested: [ [ 2, 3 ], { v: 4 } ] } ]"
    AssertEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub mapping_mixed_types()
    On Error GoTo TestFail
    Dim Globals As ASF_Globals
    GetResult "a = [1,'x',[2,'y',[3]]];" & _
                        "b = a.map(fun(x){if (IsArray(x)) {return x} elseif (IsNumeric(x)) {return x*3} else {return x}};);" & _
                        "print(b);", True
    Set Globals = scriptEngine.GetGlobals
    With Globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.Count))
    End With
    expected = "PRINT:[ 3, 'x', [ 6, 'y', [ 9 ] ] ]"
    AssertEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub filter_simple()
    On Error GoTo TestFail
    Dim Globals As ASF_Globals
    GetResult "a = [1,2,3,4];" & _
                        "b = a.filter(fun(x){ return x % 2 == 0 });" & _
                        "print(b);", True
    Set Globals = scriptEngine.GetGlobals
    With Globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.Count))
    End With
    expected = "PRINT:[ 2, 4 ]"
    AssertEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub filter_nested_arrays()
    On Error GoTo TestFail
    Dim Globals As ASF_Globals
    GetResult "a=[1,[2,3],4,[5]];" & _
                        "b=a.filter(fun(x){ return IsArray(x) });" & _
                        "print(b);", True
    Set Globals = scriptEngine.GetGlobals
    With Globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.Count))
    End With
    expected = "PRINT:[ [ 2, 3 ], [ 5 ] ]"
    AssertEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub reduce_with_initial()
    On Error GoTo TestFail
    
    actual = CStr(GetResult("a=[1,2,3,4]; return(a.reduce(fun(acc,x){ return acc + x }, 0));"))
    expected = "10"
    AssertEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub reduce_with_NO_initial()
    On Error GoTo TestFail
    
    actual = CStr(GetResult("a=[1,2,3]; return(a.reduce(fun(acc,x){ return acc + x }));"))
    expected = "6"
    AssertEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub slice_tail_only()
    On Error GoTo TestFail
    Dim Globals As ASF_Globals
    GetResult "a=[10,20,30,40]; b=a.slice(2); print(b);", True
    Set Globals = scriptEngine.GetGlobals
    With Globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.Count))
    End With
    expected = "PRINT:[ 20, 30, 40 ]"
    AssertEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub slice_start_end()
    On Error GoTo TestFail
    Dim Globals As ASF_Globals
    GetResult "a=['ant', 'bison', 'camel', 'duck', 'elephant']; b=a.slice(3,5); print(b);", True
    Set Globals = scriptEngine.GetGlobals
    With Globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.Count))
    End With
    expected = "PRINT:[ 'camel', 'duck' ]"
    AssertEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub pop_push()
    On Error GoTo TestFail
    Dim Globals As ASF_Globals
    GetResult "a=[1,2]; a.push(3); a.push(4); x = a.pop(); print(a); print(x);", True
    Set Globals = scriptEngine.GetGlobals
    With Globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.Count - 1)) & ", " _
                    & CStr(.gRuntimeLog(.gRuntimeLog.Count))
    End With
    expected = "PRINT:[ 1, 2, 3 ], PRINT:4"
    AssertEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub push_nested_member()
    On Error GoTo TestFail
    Dim Globals As ASF_Globals
    GetResult "o = { a: [1, 2], b: { c: [5] } }; o.a.push(3); o.b.c.push(6); print(o.a); print(o.b.c);", True
    Set Globals = scriptEngine.GetGlobals
    With Globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.Count - 1)) & ", " _
                    & CStr(.gRuntimeLog(.gRuntimeLog.Count))
    End With
    expected = "PRINT:[ 1, 2, 3 ], PRINT:[ 5, 6 ]"
    AssertEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub range_default()
    On Error GoTo TestFail
    Dim Globals As ASF_Globals
    GetResult "print(range(3));", True
    Set Globals = scriptEngine.GetGlobals
    With Globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.Count))
    End With
    expected = "PRINT:[ 0, 1, 2 ]"
    AssertEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub range_custom()
    On Error GoTo TestFail
    Dim Globals As ASF_Globals
    GetResult "print(range(1,3));", True
    Set Globals = scriptEngine.GetGlobals
    With Globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.Count))
    End With
    expected = "PRINT:[ 1, 2 ]"
    AssertEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub range_with_step()
    On Error GoTo TestFail
    Dim Globals As ASF_Globals
    GetResult "print(range(1,10,2));", True
    Set Globals = scriptEngine.GetGlobals
    With Globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.Count))
    End With
    expected = "PRINT:[ 1, 3, 5, 7, 9 ]"
    AssertEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub flatten_full()
    On Error GoTo TestFail
    Dim Globals As ASF_Globals
    GetResult "a=[1,[2,3],[4,[5]]]; b = flatten(a); print(b);", True
    Set Globals = scriptEngine.GetGlobals
    With Globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.Count))
    End With
    expected = "PRINT:[ 1, 2, 3, 4, 5 ]"
    AssertEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub flatten_depth_one()
    On Error GoTo TestFail
    Dim Globals As ASF_Globals
    GetResult "a=[1,[2,[3]]]; b = flatten(a,1); print(b);", True
    Set Globals = scriptEngine.GetGlobals
    With Globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.Count))
    End With
    expected = "PRINT:[ 1, 2, [ 3 ] ]"
    AssertEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub clone_array()
    On Error GoTo TestFail
    Dim Globals As ASF_Globals
    GetResult "o = { x: 1, a: [1,2] }; c = clone(o); c.a.push(3); print(o.a); print(c.a);", True
    Set Globals = scriptEngine.GetGlobals
    With Globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.Count - 1)) & ", " _
                    & CStr(.gRuntimeLog(.gRuntimeLog.Count))
    End With
    expected = "PRINT:[ 1, 2 ], PRINT:[ 1, 2, 3 ]"
    AssertEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub filter_reduce_chain()
    On Error GoTo TestFail
    
    actual = CStr(GetResult("a=[1,2,3,4,5]; return(a.filter(fun(x){ return x > 2 }).reduce(fun(acc,x){ return acc + x }, 0));"))
    expected = "12"
    AssertEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub foreach_method_updates_outer()
    On Error GoTo TestFail
    Dim Globals As ASF_Globals
    actual = CStr(GetResult("s = 0; a = [1,2,3]; a.forEach(fun(x){ s = s + x }); return(s);"))
    expected = "6"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub variables_injection()
    On Error GoTo TestFail
    Dim engine As ASF
    Dim pidx As Long

    Set engine = New ASF
    With engine
        .InjectVariable "a", Array(1, 2, 3)
        pidx = .Compile("s = 0; a.forEach(fun(x){ s = s + x }); return(s);")
        .Run pidx
        actual = CStr(.OUTPUT_)
    End With
    expected = "6"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub variables_injection_collection()
    On Error GoTo TestFail
    Dim engine As ASF
    Dim pidx As Long
    Dim coll As New Collection
    
    Set engine = New ASF
    coll.Add "We ": coll.Add "can do ": coll.Add "more with ": coll.Add "ASF!"
    With engine
        .InjectVariable "a", coll
        pidx = .Compile("txt = ''; a.forEach(fun(x){ txt += x }); return(txt);")
        .Run pidx
        actual = CStr(.OUTPUT_)
    End With
    expected = "We can do more with ASF!"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub foreach_passes_index_and_array()
    On Error GoTo TestFail
    actual = CStr(GetResult("a=[10,20]; sums=''; a.forEach(fun(v,i,arr){ sums = sums & v & ':' & i & ';' }); return(sums);"))
    expected = "10:1;20:2;"
    ' note: index semantics depend on __option_base; adjust expected if your option base is 0
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub foreach_builtin_signature_arrays()
    On Error GoTo TestFail
    actual = CStr(GetResult("s=0; foreach([1,2,3], fun(x){ s = s + x }); return(s);"))
    expected = "6"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub foreach_builtin_signature_objects()
    On Error GoTo TestFail
    actual = CStr(GetResult("s=0; c=0; foreach({math: 85, english: 92, science: 78}, fun(val, key){ s = s + val; c += 1 }); return('Average: ' + s/c);"))
    expected = "Average: 85"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub array_unique_basic()
    On Error GoTo TestFail
    Dim Globals As ASF_Globals
    GetResult "a = [1,2,2,3]; b = a.unique(); print(b);", True
    Set Globals = scriptEngine.GetGlobals
    With Globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.Count))
    End With
    expected = "PRINT:[ 1, 2, 3 ]"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub array_unique_nested()
    On Error GoTo TestFail
    Dim Globals As ASF_Globals
    GetResult "a=[1,[2],[2]]; print(a.unique());", True
    Set Globals = scriptEngine.GetGlobals
    With Globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.Count))
    End With
    expected = "PRINT:[ 1, [ 2 ] ]"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub array_concat()
    On Error GoTo TestFail
    Dim Globals As ASF_Globals
    GetResult "a=[1]; b = a.concat([2,3],4); print(b);", True
    Set Globals = scriptEngine.GetGlobals
    With Globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.Count))
    End With
    expected = "PRINT:[ 1, 2, 3, 4 ]"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub array_join_toString()
    On Error GoTo TestFail
    actual = CStr(GetResult("return(['a','b',{c:1, d:2}].join(' - '));"))
    expected = "a - b - { c: 1, d: 2 }"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub array_shift_unshift()
    On Error GoTo TestFail
    Dim Globals As ASF_Globals
    GetResult "a=[1,2,3]; x = a.shift(); a.unshift(0); print(a); print(x);", True
    Set Globals = scriptEngine.GetGlobals
    With Globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.Count - 1)) & ", " & CStr(.gRuntimeLog(.gRuntimeLog.Count))
    End With
    expected = "PRINT:[ 0, 2, 3 ], PRINT:1"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub array_delete()
    On Error GoTo TestFail
    Dim Globals As ASF_Globals
    GetResult "a=[1,2,3]; a.delete(2); print(a);", True
    Set Globals = scriptEngine.GetGlobals
    With Globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.Count))
    End With
    expected = "PRINT:[ 1, 3 ]"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub array_splice_mutating()
    On Error GoTo TestFail
    Dim Globals As ASF_Globals
    GetResult "a=[1,2,3,4]; removed=a.splice(2,2,9,10); print(removed); print(a);", True
    Set Globals = scriptEngine.GetGlobals
    With Globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.Count - 1)) & ", " & CStr(.gRuntimeLog(.gRuntimeLog.Count))
    End With
    expected = "PRINT:[ 2, 3 ], PRINT:[ 1, 9, 10, 4 ]"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Descriptionn
    Resume TestExit
End Sub

Private Sub array_toSpliced_non_mutating()
    On Error GoTo TestFail
    Dim Globals As ASF_Globals
    GetResult "a=[1,2,3]; b = a.toSpliced(2,1,9); print(a); print(b);", True
    Set Globals = scriptEngine.GetGlobals
    With Globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.Count - 1)) & ", " & CStr(.gRuntimeLog(.gRuntimeLog.Count))
    End With
    expected = "PRINT:[ 1, 2, 3 ], PRINT:[ 1, 9, 3 ]"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub array_at_negative()
    On Error GoTo TestFail
    actual = CStr(GetResult("return([10,20,30].at(-1));"))
    expected = "30"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub array_copyWithin()
    On Error GoTo TestFail
    Dim Globals As ASF_Globals
    GetResult "a=[1,2,3,4]; a.copyWithin(2,1,3); print(a);", True
    Set Globals = scriptEngine.GetGlobals
    With Globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.Count))
    End With
    expected = "PRINT:[ 1, 1, 2, 4 ]"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub array_entries()
    On Error GoTo TestFail
    Dim Globals As ASF_Globals
    GetResult "a=[10,20]; print(a.entries());", True
    Set Globals = scriptEngine.GetGlobals
    With Globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.Count))
    End With
    expected = "PRINT:[ [ 1, 10 ], [ 2, 20 ] ]"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub array_every()
    On Error GoTo TestFail
    'Check for even numbers
    actual = CStr(GetResult("return([2,4,6].every(fun(x){ return x % 2 == 0 }));"))
    expected = "True"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub array_find_and_indexes()
    On Error GoTo TestFail
    actual = CStr(GetResult("a=[1,2,3,2]; v = a.find(fun(x){ return x==2 });" & _
                            "i1 = a.findIndex(fun(x){ return x==2 });" & _
                            "i2 = a.findLastIndex(fun(x){ return x==2 });" & _
                            "v2 = a.findLast(fun(x){ return x==2 });" & _
                            "return(v & '|' & i1 & '|' & i2 & '|' & v2);"))
    expected = "2|2|4|2"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub array_from_string()
    On Error GoTo TestFail
    Dim Globals As ASF_Globals
    GetResult "print([].from('ab'))", True
    Set Globals = scriptEngine.GetGlobals
    With Globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.Count))
    End With
    expected = "PRINT:[ 'a', 'b' ]"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub from_array_copy()
    On Error GoTo TestFail
    Dim Globals As ASF_Globals
    GetResult "print([].from([1,2,3]));", True
    Set Globals = scriptEngine.GetGlobals
    With Globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.Count))
    End With
    expected = "PRINT:[ 1, 2, 3 ]"
    AssertEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub from_single_value_wrap()
    On Error GoTo TestFail
    Dim Globals As ASF_Globals
    GetResult "print([].from(5));", True
    Set Globals = scriptEngine.GetGlobals
    With Globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.Count))
    End With
    expected = "PRINT:[ 5 ]"
    AssertEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub from_with_map_array()
    On Error GoTo TestFail
    Dim Globals As ASF_Globals
    GetResult "print([].from([1,2,3], fun(x){ return x * 2 }));", True
    Set Globals = scriptEngine.GetGlobals
    With Globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.Count))
    End With
    expected = "PRINT:[ 2, 4, 6 ]"
    AssertEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub from_with_map_string()
    On Error GoTo TestFail
    Dim Globals As ASF_Globals
    GetResult "print([].from('ab', fun(c){ return c & c }));", True
    Set Globals = scriptEngine.GetGlobals
    With Globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.Count))
    End With
    expected = "PRINT:[ 'aa', 'bb' ]"
    AssertEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub from_nonclosure_second_arg_ignored()
    On Error GoTo TestFail
    Dim Globals As ASF_Globals
    ' second argument is numeric -> should be ignored and copy preserved
    GetResult "print([].from([7,8], 123));", True
    Set Globals = scriptEngine.GetGlobals
    With Globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.Count))
    End With
    expected = "PRINT:[ 7, 8 ]"
    AssertEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub array_includes_indexOf_lastIndexOf()
    On Error GoTo TestFail
    actual = CStr(GetResult("a=[1,2,3,2]; inc = a.includes(2); idx = a.indexOf(2);" & _
                            "lidx = a.lastIndexOf(2); return(inc & '|' & idx & '|' & lidx);"))
    expected = "True|2|4"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub array_of_factory_and_access()
    On Error GoTo TestFail
    actual = CStr(GetResult("return([].of(1,2,3)[2]);"))
    expected = "2"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub array_reverse_and_toReversed()
    On Error GoTo TestFail
    Dim Globals As ASF_Globals
    GetResult "a=[1,2,3]; b = a.toReversed(); a.reverse(); print(b); print(a);", True
    Set Globals = scriptEngine.GetGlobals
    With Globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.Count - 1)) & ", " & CStr(.gRuntimeLog(.gRuntimeLog.Count))
    End With
    expected = "PRINT:[ 3, 2, 1 ], PRINT:[ 3, 2, 1 ]"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub array_some()
    On Error GoTo TestFail
    actual = CStr(GetResult("return([1,3,4].some(fun(x){ return x % 2 == 0 }));"))
    expected = "True"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub array_with_non_mutating()
    On Error GoTo TestFail
    Dim Globals As ASF_Globals
    GetResult "a=[1,2,3]; b = a.with(2,9); print(a); print(b);", True
    Set Globals = scriptEngine.GetGlobals
    With Globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.Count - 1)) & ", " & CStr(.gRuntimeLog(.gRuntimeLog.Count))
    End With
    expected = "PRINT:[ 1, 2, 3 ], PRINT:[ 1, 9, 3 ]"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub array_sort_and_toSorted()
    On Error GoTo TestFail
    Dim Globals As ASF_Globals
    GetResult "a=[3,1,2]; b = a.toSorted(); a.sort(); print(b); print(a);", True
    Set Globals = scriptEngine.GetGlobals
    With Globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.Count - 1)) & ", " & CStr(.gRuntimeLog(.gRuntimeLog.Count))
    End With
    expected = "PRINT:[ 1, 2, 3 ], PRINT:[ 1, 2, 3 ]"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub array_toSpliced_and_join()
    On Error GoTo TestFail
    Dim Globals As ASF_Globals
    GetResult "a=[ 'camel', 'duck', 'elephant' ]; b = a.toSpliced(2,1,'hippo'); print(a); print(b); print(b.join(', '));", True
    Set Globals = scriptEngine.GetGlobals
    With Globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.Count - 2)) & ", " & _
                CStr(.gRuntimeLog(.gRuntimeLog.Count - 1)) & ", " & _
                CStr(.gRuntimeLog(.gRuntimeLog.Count))
    End With
    expected = "PRINT:[ 'camel', 'duck', 'elephant' ], PRINT:[ 'camel', 'hippo', 'elephant' ], PRINT:'camel, hippo, elephant'"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub


Private Sub array_entries_and_every_find_combo()
    On Error GoTo TestFail
    Dim Globals As ASF_Globals
    GetResult "a=[2,4,6]; ok = a.every(fun(x){ return x % 2 == 0 });" & _
                "f = a.find(fun(x){ return x > 4 }); print(ok); print(f);", True
    Set Globals = scriptEngine.GetGlobals
    With Globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.Count - 1)) & ", " & CStr(.gRuntimeLog(.gRuntimeLog.Count))
    End With
    expected = "PRINT:True, PRINT:6"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub array_includes_and_index_checks_with_objects()
    On Error GoTo TestFail
    Dim Globals As ASF_Globals
    GetResult "a = [{k:1},{k:2},{k:1}]; idx = a.indexOf({k:1}); inc = a.includes({k:1}); print(idx); print(inc);", True
    Set Globals = scriptEngine.GetGlobals
    With Globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.Count - 1)) & ", " & CStr(.gRuntimeLog(.gRuntimeLog.Count))
    End With
    ' indexOf returns first occurrence using deep equality; option base is 1
    expected = "PRINT:1, PRINT:True"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub array_complex_splice_and_copyWithin()
    On Error GoTo TestFail
    Dim Globals As ASF_Globals
    GetResult "a=[0,1,2,3,4,5]; removed = a.splice(3,2,9); print(removed); print(a); a.copyWithin(2,1,3); print(a);", True
    Set Globals = scriptEngine.GetGlobals
    With Globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.Count - 2)) & ", " & _
                CStr(.gRuntimeLog(.gRuntimeLog.Count - 1)) & ", " & CStr(.gRuntimeLog(.gRuntimeLog.Count))
    End With
    expected = "PRINT:[ 2, 3 ], PRINT:[ 0, 1, 9, 4, 5 ], PRINT:[ 0, 0, 1, 4, 5 ]"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub deeply_nested_Arrays()
    On Error GoTo TestFail
    Dim Globals As ASF_Globals
    GetResult "a=[1,[[2,3],4],5]; a[2][1][2]=10; print(a);", True
    Set Globals = scriptEngine.GetGlobals
    With Globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.Count))
    End With
    expected = "PRINT:[ 1, [ [ 2, 10 ], 4 ], 5 ]"
    AssertEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub
Private Sub typeof_array()
    On Error GoTo TestFail
    Dim Globals As ASF_Globals
    GetResult "a=['fruits', 'animals']; print(typeof a); print(typeof a[1]);", True
    Set Globals = scriptEngine.GetGlobals
    With Globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.Count - 1)) & ", " _
                & CStr(.gRuntimeLog(.gRuntimeLog.Count))
    End With
    expected = "PRINT:'array', PRINT:'string'"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub typeof_object()
    On Error GoTo TestFail
    Dim Globals As ASF_Globals
    GetResult "a={name: 'mango', fruit: true}; print(typeof a); print(typeof a.fruit);", True
    Set Globals = scriptEngine.GetGlobals
    With Globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.Count - 1)) & ", " _
                & CStr(.gRuntimeLog(.gRuntimeLog.Count))
    End With
    expected = "PRINT:'object: <ASF_Map>', PRINT:'boolean'"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub typeof_closures()
    On Error GoTo TestFail
    Dim Globals As ASF_Globals
    GetResult "f = fun(a) { a = a + 1; return a }; print(typeof f); print(typeof f(1));", True
    Set Globals = scriptEngine.GetGlobals
    With Globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.Count - 1)) & ", " _
                & CStr(.gRuntimeLog(.gRuntimeLog.Count))
    End With
    expected = "PRINT:'function', PRINT:'number'"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub for_of_with_array()
    On Error GoTo TestFail
    actual = CStr(GetResult("a=[1,2,3]; s=0; for (v of a) { s = s + v }; return(s);"))
    expected = "6"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub for_in_with_array_indices()
    On Error GoTo TestFail
    Dim Globals As ASF_Globals
    GetResult "a=[10,20]; out=[]; for (i in a) { out.push(i) }; print(out);", True
    Set Globals = scriptEngine.GetGlobals
    With Globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.Count))
    End With
    expected = "PRINT:[ 1, 2 ]"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub for_of_with_string()
    On Error GoTo TestFail
    actual = CStr(GetResult("s='ab'; out=''; for (ch of s) { out = out + ch }; return(out);"))
    expected = "ab"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub for_in_with_object_properties()
    On Error GoTo TestFail
    Dim Globals As ASF_Globals
    GetResult "o = { a:1, b:2 }; keys=[]; for (k in o) { keys.push(k) }; print(keys);", True
    Set Globals = scriptEngine.GetGlobals
    With Globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.Count))
    End With
    expected = "PRINT:[ 'a', 'b' ]"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub string_at()
    On Error GoTo TestFail
    actual = CStr(GetResult("b='ABCD'.at(1); return(b);"))
    expected = "B"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub string_at_negative_index()
    On Error GoTo TestFail
    actual = CStr(GetResult("b='ABCD'.at(-2); return(b);"))
    expected = "C"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub string_at_default_index()
    On Error GoTo TestFail
    actual = CStr(GetResult("b='ABCD'.at; return(b);"))
    expected = "A"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub string_length()
    On Error GoTo TestFail
    actual = CStr(GetResult("b='ABCD'.length(); return(b);"))
    expected = "4"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub string_charAt()
    On Error GoTo TestFail
    actual = CStr(GetResult("b='ABCD'; a=b.charAt(b.length - 1); return(a);"))
    expected = "D"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub string_charCodeAt()
    On Error GoTo TestFail
    actual = CStr(GetResult("b='ABCD'; a=b.charCodeAt(2); return(a);"))
    expected = "67"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub string_concat()
    On Error GoTo TestFail
    actual = CStr(GetResult("a='ABCD'; b='EFGH'; c='IJKL'; d=a.concat(' + ', b, c); return(d);"))
    expected = "ABCD + EFGH + IJKL"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub string_endsWith()
    On Error GoTo TestFail
    actual = CBool(GetResult("a='Scripting for all'; b=a.endsWith('all'); return(b);"))
    expected = True
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub string_fromCharCode()
    On Error GoTo TestFail
    actual = CStr(GetResult("a=''.fromCharCode(68); return(a);"))
    expected = "D"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub string_includes()
    On Error GoTo TestFail
    actual = CBool(GetResult("a='Bridging the gap with modern programming'.includes('modern'); return(a);"))
    expected = True
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub string_indexOf()
    On Error GoTo TestFail
    actual = CStr(GetResult("a='Bridging the gap with modern programming'.indexOf('gap'); return(a);"))
    expected = "13"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub string_lastIndexOf()
    On Error GoTo TestFail
    actual = CStr(GetResult("a='Flow control, functions, closures'.lastIndexOf(',', 8); return(a);"))
    expected = "23"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub string_localCompare()
    On Error GoTo TestFail
    actual = CStr(GetResult("a='AB'.localeCompare('ab'); return(a);"))
    expected = "-1"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub string_pad()
    On Error GoTo TestFail
    actual = CStr(GetResult("a='AB'.padEnd(9, '+_'); return(a.padStart(17, '+_'));"))
    expected = "+_+_+_+_AB+_+_+_+"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub string_template()
    On Error GoTo TestFail
    actual = CStr(GetResult("a='Happy! '; return(`I feel ${a.repeat(3)}`);"))
    expected = "I feel Happy! Happy! Happy! "
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub string_template_with_arrays()
    On Error GoTo TestFail
    actual = CStr(GetResult("a=[1,2]; return(`arr:${a[1] + a[2]} sum two items`);"))
    expected = "arr:3 sum two items"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub string_replace()
    On Error GoTo TestFail
    actual = CStr(GetResult("welcome=fun(string){return string.concat('!')};" & _
                    "return('Hello world'.replace('world', welcome('VBA')));"))
    expected = "Hello VBA!"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub string_slice()
    On Error GoTo TestFail
    actual = CStr(GetResult("return('Bridging old VBA to modern syntax.'.slice(-21, -1));"))
    expected = "VBA to modern syntax"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub string_splite()
    On Error GoTo TestFail
    actual = CStr(GetResult("chars='Bridging old VBA to modern syntax.'.split('', 4); return(chars[4]);"))
    expected = "d"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub string_startWith()
    On Error GoTo TestFail
    actual = CBool(GetResult("a='Scripting for all'; b=a.startsWith('Script'); return(b);"))
    expected = True
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub string_subString()
    On Error GoTo TestFail
    actual = CStr(GetResult("return('Do more inside VBA'.substring(0,3));"))
    expected = "Do "
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub string_toLowercase()
    On Error GoTo TestFail
    actual = CStr(GetResult("return('Do more inside VBA'.toLowercase);"))
    expected = "do more inside vba"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub string_toUppercase()
    On Error GoTo TestFail
    actual = CStr(GetResult("return('Do more inside VBA'.toUppercase);"))
    expected = "DO MORE INSIDE VBA"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub string_trim()
    On Error GoTo TestFail
    actual = CStr(GetResult("return('   less boilerplate    '.trim);"))
    expected = "less boilerplate"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub string_trim_start_end()
    On Error GoTo TestFail
    actual = CStr(GetResult("return('   less boilerplate    '.trimStart().trimEnd);"))
    expected = "less boilerplate"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub regex_replace_from_string_object()
    On Error GoTo TestFail
    actual = CStr(GetResult("return('I think my Dog is cuter than your dog!'.replace(`/dog/i`, 'cat'));"))
    expected = "I think my cat is cuter than your dog!"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub regex_replacer_function_from_string_object()
    On Error GoTo TestFail
    actual = CStr(GetResult("fun replacer(match, p1, p2, p3, offset, string)" & _
                                "{return [p1, p2, p3].join(' - ');};" & _
                            "newString = 'abc12345#$*%'.replace(`/(\D*)(\d*)(\W*)/`, replacer);" & _
                            "return(newString);"))
    expected = "abc - 12345 - #$*%"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub regex_replace_using_placeholders_from_string_object()
    On Error GoTo TestFail
    actual = CStr(GetResult("return('Maria Cruz'.replace(`/(\w+)\s(\w+)/`, '$2, $1'));"))
    expected = "Cruz, Maria"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub regex_editing_matches_from_string_object()
    On Error GoTo TestFail
    actual = CStr(GetResult("fun styleHyphenFormat(propertyName) {" & _
                                "upperToHyphenLower = fun(match, offset, string) {" & _
                                    "return (offset > 0 ? ' - ' : '') + match.toLowerCase();" & _
                                "};" & _
                            "return propertyName.replace(`/[A-Z]/g`, upperToHyphenLower);" & _
                            "}; return(styleHyphenFormat('borderTop'));"))
    expected = "border - top"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub regex_replace_using_templates_from_string_object()
    On Error GoTo TestFail
    actual = CStr(GetResult("fun superSafeRedactName(text, name) {" & _
                                "return text.replaceAll(`/${regex().escape(name)}/g`, '[REDACTED]');" & _
                            "};" & _
                            "report = 'A hacker called acke breached the system.';" & _
                            "return(superSafeRedactName(report, 'acke')); /* 'A h[REDACTED]r called [REDACTED] breached the system.'*/"))
    expected = "A h[REDACTED]r called [REDACTED] breached the system."
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub regex_match_from_string_object()
    On Error GoTo TestFail
    Dim Globals As ASF_Globals
    GetResult "print('test1test2'.match(`/t(e)(st(\d?))/`));", True
    Set Globals = scriptEngine.GetGlobals
    With Globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.Count))
    End With
    expected = "PRINT:[ 'test1', 'e', 'st1', '1' ]"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub regex_match_global_from_string_object()
    On Error GoTo TestFail
    Dim Globals As ASF_Globals
    GetResult "print('test1test2'.match(`/t(e)(st(\d?))/g`));", True
    Set Globals = scriptEngine.GetGlobals
    With Globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.Count))
    End With
    expected = "PRINT:[ 'test1', 'test2' ]"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub regex_matchall_from_string_object()
    On Error GoTo TestFail
    Dim Globals As ASF_Globals
    GetResult "print('test1test2'.matchAll(`/t(e)(st(\d?))/g`));", True
    Set Globals = scriptEngine.GetGlobals
    With Globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.Count))
    End With
    expected = "PRINT:[ [ 'test1', 'e', 'st1', '1' ], [ 'test2', 'e', 'st2', '2' ] ]"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub regex_replace_empty_alternative_global()
    On Error GoTo TestFail
    actual = CStr(GetResult("return('999'.replace(`/(9|)/g`, 'x'));"))
    expected = "xxxx"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub regex_replace_empty_match_keeps_chars()
    On Error GoTo TestFail
    actual = CStr(GetResult("return('a9'.replace(`/(9|)/g`, 'x'));"))
    expected = "xaxx"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub regex_replace_lookahead_insertion()
    On Error GoTo TestFail
    actual = CStr(GetResult("return('ab'.replace(`/(?=b)/`, 'x'));"))
    expected = "axb"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub regex_replace_empty_match_at_start()
    On Error GoTo TestFail
    actual = CStr(GetResult("return('999'.replace(`/x*/`, 'y'));"))
    expected = "y999"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub regex_match_global_empty_alternative()
    On Error GoTo TestFail
    Dim Globals As ASF_Globals
    GetResult "print('999'.match(`/9|/g`));", True
    Set Globals = scriptEngine.GetGlobals
    With Globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.Count))
    End With
    expected = "PRINT:[ '9', '9', '9', '' ]"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub regex_replaceall_requires_global_flag()
    On Error GoTo TestFail
    actual = CStr(GetResult("try { 'a-b-c'.replaceAll(`/-/`, '+'); return('no error') }" & _
                            "catch { return('caught') }"))
    expected = "caught"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub regex_replaceall_with_global_flag()
    On Error GoTo TestFail
    actual = CStr(GetResult("return('a-b-c'.replaceAll(`/-/g`, '+'));"))
    expected = "a+b+c"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub regex_object_constructor()
    On Error GoTo TestFail
    actual = CStr(GetResult("re=regex(); return(typeOf(re));"))
    expected = "object: <ASF_RegexEngine>"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub regex_object_execute_method()
    On Error GoTo TestFail
    Dim Globals As ASF_Globals
    GetResult "re=regex(); re.init(`(\D*)(\d*)(\W*)`); print(re.exec('abc12345#$*%'));", True
    Set Globals = scriptEngine.GetGlobals
    With Globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.Count))
    End With
    expected = "PRINT:[ 'abc12345#$*%', 'abc', '12345', '#$*%' ]"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub regex_object_replace_method()
    On Error GoTo TestFail
    actual = CStr(GetResult("re=regex(); re.init(`(foo)(bar)`); return(re.replace('foobar', '$2-$1'));"))
    expected = "bar-foo"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub regex_object_split_method()
    On Error GoTo TestFail
    Dim Globals As ASF_Globals
    GetResult "re=regex(); re.init(`[,;\.\s]+`); print(re.split('apple,orange;banana grape.strawberry'));", True
    Set Globals = scriptEngine.GetGlobals
    With Globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.Count))
    End With
    expected = "PRINT:[ 'apple', 'orange', 'banana', 'grape', 'strawberry' ]"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub regex_object_executeAll_method()
    On Error GoTo TestFail
    Dim Globals As ASF_Globals
    GetResult "re=regex(`[,;\.\s]+`); print(re.ExecAll('apple,orange;banana grape.strawberry'));", True
    Set Globals = scriptEngine.GetGlobals
    With Globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.Count))
    End With
    expected = "PRINT:[ [ ',' ], [ ';' ], [ ' ' ], [ '.' ] ]"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub regex_object_method_with_flag_property()
    On Error GoTo TestFail
    actual = CBool(GetResult("re=regex(`[a-z]`); re.setignorecase(False); return(re.Test('A'));"))
    expected = False
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub regex_object_named_captures()
    On Error GoTo TestFail
    actual = CStr(GetResult("re=regex(`(?<nonDigits>\D*)(?<digits>\d+)(?<nonWords>\W*)`);" _
                            & "return(re.Replace('abc123#$', '[$<nonWords>]|[$<digits>]|[$<nonDigits>]'));"))
    expected = "[#$]|[123]|[abc]"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub regex_object_conditional_simple()
    On Error GoTo TestFail
    Dim Globals As ASF_Globals
    GetResult "re=regex(`(?:(a)|(b))(?(1)(X)|(Y))`); print(re.Exec('aX'));", True
    Set Globals = scriptEngine.GetGlobals
    With Globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.Count))
    End With
    expected = "PRINT:[ 'aX', 'a', '', 'X', '' ]"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub regex_object_conditional_inline_flags()
    On Error GoTo TestFail
    Dim Globals As ASF_Globals
    GetResult "re=regex(`(?:(a)|(b))(?(1)(?i:x)|(?i:y))`); print(re.Exec('aaaX'));", True
    Set Globals = scriptEngine.GetGlobals
    With Globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.Count))
    End With
    expected = "PRINT:[ 'aX', 'a', '' ]"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub regex_object_conditional_inline_flags_off()
    On Error GoTo TestFail
    Dim Globals As ASF_Globals
    GetResult "re=regex(`(?:(a)|(b))(?(1)(?-i:x)|y)`, True); print(re.Exec('bbBY'));", True
    Set Globals = scriptEngine.GetGlobals
    With Globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.Count))
    End With
    expected = "PRINT:[ 'BY', '', 'B' ]"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub regex_object_inline_dotAll_flags()
    On Error GoTo TestFail
    Dim Globals As ASF_Globals
    GetResult "re=regex(`(?:(a)|(b))(?(1)(?s:.))`); print(re.Exec('a\n'));", True
    Set Globals = scriptEngine.GetGlobals
    With Globals
        actual = ConvertNewLines(CStr(.gRuntimeLog(.gRuntimeLog.Count)))
    End With
    expected = "PRINT:[ 'a\n', 'a', '' ]"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub regex_object_inline_dotAll_flags_off()
    On Error GoTo TestFail
    actual = CStr(GetResult("re=regex(`(?:(a)|(b))(?(1)(?-s:.))`); return(re.Exec('a\n'));"))
    expected = vbNullString
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub regex_object_Named_Conditionals()
    On Error GoTo TestFail
    Dim Globals As ASF_Globals
    GetResult "re=regex(`(?:(?<A>a)|(?<B>b))(?(A)X|Y)`); print(re.Exec('bY'));", True
    Set Globals = scriptEngine.GetGlobals
    With Globals
        actual = ConvertNewLines(CStr(.gRuntimeLog(.gRuntimeLog.Count)))
    End With
    expected = "PRINT:[ 'bY', '', 'b' ]"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub regex_object_Named_Conditionals_NO_Match()
    On Error GoTo TestFail
    actual = CStr(GetResult("re=regex(`(?:(?<A>a)|(?<B>b))(?(A)X|Y)`); print(re.Exec('aY'));"))
    expected = vbNullString
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub classes_static_method_declarations()
    On Error GoTo TestFail
    Dim Globals As ASF_Globals
    GetResult "class A {" & vbCrLf & _
            "  field x = 1, y = 2;" & vbCrLf & _
            "  foo() { return this.x + 1 };" & vbCrLf & _
            "  static greet() { return 'hello' };" & vbCrLf & _
            "};" & vbCrLf & _
            "a = new A();" & vbCrLf & _
            "r1 = a.foo();" & vbCrLf & _
            "r2 = A.greet();" & vbCrLf & _
            "print(r1);" & vbCrLf & _
            "print(r2);", True
    Set Globals = scriptEngine.GetGlobals
    With Globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.Count - 1)) & ", " _
                & CStr(.gRuntimeLog(.gRuntimeLog.Count))
    End With
    expected = "PRINT:2, PRINT:'hello'"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub classes_field_declarations()
    On Error GoTo TestFail
    Dim Globals As ASF_Globals
    GetResult "class Animal {" & vbCrLf & _
            "  field name, age = 0;" & vbCrLf & _
            "  constructor(n) {" & vbCrLf & _
            "    this.name = n;" & vbCrLf & _
            "  };" & vbCrLf & _
            "  speak() {" & vbCrLf & _
            "    print('Animal ' + this.name + ' is ' + this.age + ' years old');" & vbCrLf & _
            "  };" & vbCrLf & _
            "};" & vbCrLf & _
            "let animal = new Animal('Lion');" & vbCrLf & _
            "animal.age = 5;" & vbCrLf & _
            "animal.speak();", True
    Set Globals = scriptEngine.GetGlobals
    With Globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.Count))
    End With
    expected = "PRINT:'Animal Lion is 5 years old'"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub classes_inheritance_with_fields()
    On Error GoTo TestFail
    Dim Globals As ASF_Globals
    GetResult "class Shape {" & vbCrLf & _
           "  field x = 0, y = 0, color = 'black';" & vbCrLf & _
           "  constructor(x, y) {" & vbCrLf & _
           "    this.x = x;" & vbCrLf & _
           "    this.y = y;" & vbCrLf & _
           "  };" & vbCrLf & _
           "  getPosition() {" & vbCrLf & _
           "    return this.x + ',' + this.y;" & vbCrLf & _
           "  };" & vbCrLf & _
           "};" & vbCrLf & _
           "class Circle extends Shape {" & vbCrLf & _
           "  field radius = 1;" & vbCrLf & _
           "  constructor(x, y, r) {" & vbCrLf & _
           "    super(x, y);" & vbCrLf & _
           "    this.radius = r;" & vbCrLf & _
           "    this.color = 'red';" & vbCrLf & _
           "  };" & vbCrLf & _
           "  getArea() {" & vbCrLf & _
           "    return 3.14159 * this.radius * this.radius;" & vbCrLf & _
           "  };" & vbCrLf & _
           "};" & vbCrLf & _
           "let circle = new Circle(10, 20, 5);" & vbCrLf & _
           "print('Position: ' + circle.getPosition());" & vbCrLf & _
           "print('Color: ' + circle.color);" & vbCrLf & _
           "print('Area: ' + circle.getArea());", True
    Set Globals = scriptEngine.GetGlobals
    With Globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.Count - 2)) & ", " _
                & CStr(.gRuntimeLog(.gRuntimeLog.Count - 1)) & ", " _
                & CStr(.gRuntimeLog(.gRuntimeLog.Count))
    End With
    expected = "PRINT:'Position: 10,20', PRINT:'Color: red', PRINT:'Area: 78.53975'"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub classes_fields_declarations_with_mixed_syntax()
    On Error GoTo TestFail
    actual = CStr(GetResult("class Counter {" & vbCrLf & _
                            "  field count = 0, step = 1, max = 100;" & vbCrLf & _
                            "  field name;" & vbCrLf & _
                            "  constructor(n) {" & vbCrLf & _
                            "    this.name = n;" & vbCrLf & _
                            "  };" & vbCrLf & _
                            "  increment() {" & vbCrLf & _
                            "    if (this.count < this.max) {" & vbCrLf & _
                            "      this.count = this.count + this.step;" & vbCrLf & _
                            "    };" & vbCrLf & _
                            "  };" & vbCrLf & _
                            "  getValue() {" & vbCrLf & _
                            "    return this.name + ': ' + this.count;" & vbCrLf & _
                            "  };" & vbCrLf & _
                            "  static create(name, step) {" & vbCrLf & _
                            "    let c = new Counter(name);" & vbCrLf & _
                            "    c.step = step;" & vbCrLf & _
                            "    return c;" & vbCrLf & _
                            "  };" & vbCrLf & _
                            "};" & vbCrLf & _
                            "let counter = Counter.create('MyCounter', 5);" & vbCrLf & _
                            "counter.increment();" & vbCrLf & _
                            "counter.increment();" & vbCrLf & _
                            "return counter.getValue();"))
    expected = "MyCounter: 10"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub classes_fields_with_complex_initializers()
    On Error GoTo TestFail
    Dim Globals As ASF_Globals
    GetResult "class Vector {" & vbCrLf & _
            "  field x = 0, y = 0, z = 0;" & vbCrLf & _
            "  field magnitude = 0;" & vbCrLf & _
            "  constructor(x, y, z) {" & vbCrLf & _
            "    this.x = x;" & vbCrLf & _
            "    this.y = y;" & vbCrLf & _
            "    this.z = z;" & vbCrLf & _
            "    this.updateMagnitude();" & vbCrLf & _
            "  };" & vbCrLf & _
            "  updateMagnitude() {" & vbCrLf & _
            "    this.magnitude = (this.x * this.x + this.y * this.y + this.z * this.z) ^ 0.5;" & vbCrLf & _
            "  };" & vbCrLf & _
            "  toString() {" & vbCrLf & _
            "    return '(' + this.x + ', ' + this.y + ', ' + this.z + ')';" & vbCrLf & _
            "  };" & vbCrLf & _
            "};" & vbCrLf & _
            "let v = new Vector(3, 4, 0);" & vbCrLf & _
            "print('Vector: ' + v.toString());" & vbCrLf & _
            "print('Magnitude: ' + v.magnitude);", True
    Set Globals = scriptEngine.GetGlobals
    With Globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.Count - 1)) & ", " _
                & CStr(.gRuntimeLog(.gRuntimeLog.Count))
    End With
    expected = "PRINT:'Vector: (3, 4, 0)', PRINT:'Magnitude: 5'"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub classes_multiple_instances()
    On Error GoTo TestFail
    Dim Globals As ASF_Globals
    GetResult "class Counter {" & vbCrLf & _
            "  field count = 0;" & vbCrLf & _
            "  increment() {" & vbCrLf & _
            "    this.count = this.count + 1;" & vbCrLf & _
            "  };" & vbCrLf & _
            "  getCount() {" & vbCrLf & _
            "    return this.count;" & vbCrLf & _
            "  };" & vbCrLf & _
            "};" & vbCrLf & _
            "let c1 = new Counter();" & vbCrLf & _
            "let c2 = new Counter();" & vbCrLf & _
            "c1.increment();" & vbCrLf & _
            "c1.increment();" & vbCrLf & _
            "c2.increment();" & vbCrLf & _
            "print('Counter 1: ' + c1.getCount());" & vbCrLf & _
            "print('Counter 2: ' + c2.getCount());", True
    Set Globals = scriptEngine.GetGlobals
    With Globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.Count - 1)) & ", " _
                & CStr(.gRuntimeLog(.gRuntimeLog.Count))
    End With
    expected = "PRINT:'Counter 1: 2', PRINT:'Counter 2: 1'"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub classes_method_chaining()
    On Error GoTo TestFail
    actual = CStr(GetResult("class Builder {" & vbCrLf & _
                            "  field value = '';" & vbCrLf & _
                            "  add(text) {" & vbCrLf & _
                            "    this.value = this.value + text;" & vbCrLf & _
                            "    return this;" & vbCrLf & _
                            "  };" & vbCrLf & _
                            "  build() {" & vbCrLf & _
                            "    return this.value;" & vbCrLf & _
                            "  };" & vbCrLf & _
                            "};" & vbCrLf & _
                            "let result = new Builder().add('Hello').add(' ').add('World').build();" & vbCrLf & _
                            "return result;"))
    expected = "Hello World"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub object_keys()
    On Error GoTo TestFail
    Dim Globals As ASF_Globals
    GetResult "o = {a: 1, b: 2, c: 3}; print(o.keys());", True
    Set Globals = scriptEngine.GetGlobals
    With Globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.Count))
    End With
    expected = "PRINT:[ 'a', 'b', 'c' ]"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub object_values()
    On Error GoTo TestFail
    Dim Globals As ASF_Globals
    GetResult "o = {a: 1, b: 2, c: 3}; print(o.values());", True
    Set Globals = scriptEngine.GetGlobals
    With Globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.Count))
    End With
    expected = "PRINT:[ 1, 2, 3 ]"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub object_entries()
    On Error GoTo TestFail
    Dim Globals As ASF_Globals
    GetResult "o = {a: 1, b: 2}; print(o.entries());", True
    Set Globals = scriptEngine.GetGlobals
    With Globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.Count))
    End With
    expected = "PRINT:[ [ 'a', 1 ], [ 'b', 2 ] ]"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub object_size()
    On Error GoTo TestFail
    actual = CStr(GetResult("o = {a: 1, b: 2, c: 3}; return(o.size());"))
    expected = "3"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub object_length()
    On Error GoTo TestFail
    actual = CStr(GetResult("o = {x: 10, y: 20}; return(o.length());"))
    expected = "2"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub object_hasKey_true()
    On Error GoTo TestFail
    actual = CBool(GetResult("o = {a: 1, b: 2}; return(o.hasKey('a'));"))
    expected = True
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub object_hasKey_false()
    On Error GoTo TestFail
    actual = CBool(GetResult("o = {a: 1, b: 2}; return(o.hasKey('c'));"))
    expected = False
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub object_has_alias()
    On Error GoTo TestFail
    actual = CBool(GetResult("o = {name: 'John'}; return(o.has('name'));"))
    expected = True
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub object_isEmpty_true()
    On Error GoTo TestFail
    actual = CBool(GetResult("o = {}; return(o.isEmpty());"))
    expected = True
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub object_isEmpty_false()
    On Error GoTo TestFail
    actual = CBool(GetResult("o = {a: 1}; return(o.isEmpty());"))
    expected = False
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub object_get_existing_key()
    On Error GoTo TestFail
    actual = CStr(GetResult("o = {a: 1, b: 2}; return(o.get('b'));"))
    expected = "2"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub object_get_with_default()
    On Error GoTo TestFail
    actual = CStr(GetResult("o = {a: 1}; return(o.get('b', 99));"))
    expected = "99"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub object_get_without_default_missing_key()
    On Error GoTo TestFail
    actual = CStr(GetResult("o = {a: 1}; return(o.get('b'));"))
    expected = vbNullString
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub object_set_new_key()
    On Error GoTo TestFail
    Dim Globals As ASF_Globals
    GetResult "o = {a: 1}; o.set('b', 2); print(o);", True
    Set Globals = scriptEngine.GetGlobals
    With Globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.Count))
    End With
    expected = "PRINT:{ a: 1, b: 2 }"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub object_set_update_existing()
    On Error GoTo TestFail
    actual = CStr(GetResult("o = {a: 1}; o.set('a', 10); return(o.a);"))
    expected = "10"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub object_delete_existing_key()
    On Error GoTo TestFail
    Dim Globals As ASF_Globals
    GetResult "o = {a: 1, b: 2, c: 3}; o.delete('b'); print(o);", True
    Set Globals = scriptEngine.GetGlobals
    With Globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.Count))
    End With
    expected = "PRINT:{ a: 1, c: 3 }"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub object_remove_alias()
    On Error GoTo TestFail
    actual = CStr(GetResult("o = {a: 1, b: 2}; o.remove('a'); return(o.size());"))
    expected = "1"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub object_clear()
    On Error GoTo TestFail
    actual = CStr(GetResult("o = {a: 1, b: 2, c: 3}; o.clear(); return(o.isEmpty());"))
    expected = "True"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub object_clone_simple()
    On Error GoTo TestFail
    Dim Globals As ASF_Globals
    GetResult "o = {a: 1, b: 2}; c = o.clone(); c.b = 99; print(o.b); print(c.b);", True
    Set Globals = scriptEngine.GetGlobals
    With Globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.Count - 1)) & ", " & _
                CStr(.gRuntimeLog(.gRuntimeLog.Count))
    End With
    expected = "PRINT:2, PRINT:99"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub object_clone_nested()
    On Error GoTo TestFail
    Dim Globals As ASF_Globals
    GetResult "o = {x: 1, nested: {y: 2}}; c = o.clone(); c.nested.y = 99; print(o.nested.y); print(c.nested.y);", True
    Set Globals = scriptEngine.GetGlobals
    With Globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.Count - 1)) & ", " & _
                CStr(.gRuntimeLog(.gRuntimeLog.Count))
    End With
    expected = "PRINT:2, PRINT:99"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub object_merge()
    On Error GoTo TestFail
    Dim Globals As ASF_Globals
    GetResult "o1 = {a: 1, b: 2}; o2 = {b: 20, c: 3}; o1.merge(o2); print(o1);", True
    Set Globals = scriptEngine.GetGlobals
    With Globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.Count))
    End With
    expected = "PRINT:{ a: 1, b: 20, c: 3 }"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub object_merge_nested()
    On Error GoTo TestFail
    Dim Globals As ASF_Globals
    GetResult "o1 = {a: 1, nested: {x: 10}}; o2 = {nested: {y: 20}, b: 2}; o1.merge(o2); print(o1);", True
    Set Globals = scriptEngine.GetGlobals
    With Globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.Count))
    End With
    expected = "PRINT:{ a: 1, nested: { y: 20 }, b: 2 }"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub object_forEach()
    On Error GoTo TestFail
    actual = CStr(GetResult("o = {a: 1, b: 2, c: 3}; s = 0; o.forEach(fun(val, key) { s = s + val }); return(s);"))
    expected = "6"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub object_forEach_with_key()
    On Error GoTo TestFail
    actual = CStr(GetResult("o = {a: 1, b: 2}; result = ''; o.forEach(fun(val, key) { result = result & key & ':' & val & ';' }); return(result);"))
    expected = "a:1;b:2;"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub object_map()
    On Error GoTo TestFail
    Dim Globals As ASF_Globals
    GetResult "o = {a: 1, b: 2, c: 3}; result = o.map(fun(val, key) { return val * 2 }); print(result);", True
    Set Globals = scriptEngine.GetGlobals
    With Globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.Count))
    End With
    expected = "PRINT:{ a: 2, b: 4, c: 6 }"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub object_map_with_key()
    On Error GoTo TestFail
    Dim Globals As ASF_Globals
    GetResult "o = {a: 1, b: 2}; result = o.map(fun(val, key) { return key & val }); print(result);", True
    Set Globals = scriptEngine.GetGlobals
    With Globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.Count))
    End With
    expected = "PRINT:{ a: 'a1', b: 'b2' }"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub object_filter()
    On Error GoTo TestFail
    Dim Globals As ASF_Globals
    GetResult "o = {a: 1, b: 2, c: 3, d: 4}; result = o.filter(fun(val, key) { return val % 2 == 0 }); print(result);", True
    Set Globals = scriptEngine.GetGlobals
    With Globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.Count))
    End With
    expected = "PRINT:{ b: 2, d: 4 }"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub object_filter_by_key()
    On Error GoTo TestFail
    Dim Globals As ASF_Globals
    GetResult "o = {apple: 10, banana: 20, cherry: 30}; result = o.filter(fun(val, key) { return key.startsWith('a') }); print(result);", True
    Set Globals = scriptEngine.GetGlobals
    With Globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.Count))
    End With
    expected = "PRINT:{ apple: 10 }"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub object_some_true()
    On Error GoTo TestFail
    actual = CBool(GetResult("o = {a: 1, b: 2, c: 3}; return(o.some(fun(val) { return val > 2 }));"))
    expected = True
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub object_some_false()
    On Error GoTo TestFail
    actual = CBool(GetResult("o = {a: 1, b: 2}; return(o.some(fun(val) { return val > 10 }));"))
    expected = False
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub object_every_true()
    On Error GoTo TestFail
    actual = CBool(GetResult("o = {a: 2, b: 4, c: 6}; return(o.every(fun(val) { return val % 2 == 0 }));"))
    expected = True
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub object_every_false()
    On Error GoTo TestFail
    actual = CBool(GetResult("o = {a: 2, b: 3, c: 4}; return(o.every(fun(val) { return val % 2 == 0 }));"))
    expected = False
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub object_chaining_methods()
    On Error GoTo TestFail
    Dim Globals As ASF_Globals
    GetResult "o = {a: 1, b: 2, c: 3, d: 4}; result = o.filter(fun(v) { return v > 1 }).map(fun(v) { return v * 10 }); print(result);", True
    Set Globals = scriptEngine.GetGlobals
    With Globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.Count))
    End With
    expected = "PRINT:{ b: 20, c: 30, d: 40 }"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub object_with_nested_arrays()
    On Error GoTo TestFail
    Dim Globals As ASF_Globals
    GetResult "o = {a: [1,2], b: [3,4]}; result = o.map(fun(arr) { return arr.reduce(fun(sum, x) { return sum + x }, 0) }); print(result);", True
    Set Globals = scriptEngine.GetGlobals
    With Globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.Count))
    End With
    expected = "PRINT:{ a: 3, b: 7 }"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub object_entries_with_forEach()
    On Error GoTo TestFail
    actual = CStr(GetResult("o = {x: 10, y: 20}; s = 0; o.entries().forEach(fun(pair) { s = s + pair[2] }); return(s);"))
    expected = "30"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub polymorphism_basic()
    On Error GoTo TestFail
    Dim Globals As ASF_Globals
    GetResult "class Animal {" & _
                "    field name;" & _
                "    constructor(n) { this.name = n; }" & _
                "    speak() { return 'Some sound'; }" & _
                "};" & _
                "class Dog extends Animal {" & _
                "    speak() { return this.name + ' barks'; }" & _
                "};" & _
                "class Cat extends Animal {" & _
                "    speak() { return this.name + ' meows'; }" & _
                "};" & _
                "dog = new Dog('Rex');" & _
                "cat = new Cat('Whiskers');" & _
                "print(dog.speak());" & _
                "print(cat.speak());", True
    Set Globals = scriptEngine.GetGlobals
    With Globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.Count - 1)) & ", " & _
                CStr(.gRuntimeLog(.gRuntimeLog.Count))
    End With
    expected = "PRINT:'Rex barks', PRINT:'Whiskers meows'"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub polymorphism_three_levels()
    On Error GoTo TestFail
    Dim Globals As ASF_Globals
    GetResult "class Vehicle {" & _
                "    move() { return 'moving'; }" & _
                "};" & _
                "class Car extends Vehicle {" & _
                "    move() { return 'driving on road'; }" & _
                "};" & _
                "class SportsCar extends Car {" & _
                "    move() { return 'racing on track'; }" & _
                "};" & _
                "v = new Vehicle();" & _
                "c = new Car();" & _
                "s = new SportsCar();" & _
                "print(v.move());" & _
                "print(c.move());" & _
                "print(s.move());", True
    Set Globals = scriptEngine.GetGlobals
    With Globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.Count - 2)) & ", " & _
                CStr(.gRuntimeLog(.gRuntimeLog.Count - 1)) & ", " & _
                CStr(.gRuntimeLog(.gRuntimeLog.Count))
    End With
    expected = "PRINT:'moving', PRINT:'driving on road', PRINT:'racing on track'"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub polymorphism_method_delegation()
    On Error GoTo TestFail
    actual = CStr(GetResult("class Printer {" & _
                "    print(doc) { return 'Printing: ' + doc; };" & _
                "};" & _
                "class ColorPrinter extends Printer {" & _
                "    print(doc) { return 'Color printing: ' + doc; };" & _
                "};" & _
                "class LaserPrinter extends Printer {" & _
                "    print(doc) { return 'Laser printing: ' + doc; };" & _
                "};" & _
                "fun printDocument(printer, doc) {" & _
                "    return printer.print(doc);" & _
                "};" & _
                "p1 = new Printer();" & _
                "p2 = new ColorPrinter();" & _
                "p3 = new LaserPrinter();" & _
                "result = [printDocument(p1, 'Doc1'), " & _
                "         printDocument(p2, 'Doc2'), " & _
                "         printDocument(p3, 'Doc3')].join(' | ');" & _
                "return result;"))
    expected = "Printing: Doc1 | Color printing: Doc2 | Laser printing: Doc3"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub sort_chain_on_objects()
    On Error GoTo TestFail
    Dim Globals As ASF_Globals
    Dim jsonResponse As String
    jsonResponse = _
        "{" & _
        "  users: [" & _
        "    { id: 1, name: 'Alice', sales: 15000, active: true }," & _
        "    { id: 2, name: 'Bob', sales: 8000, active: false }," & _
        "    { id: 3, name: 'Charlie', sales: 22000, active: true }" & _
        "  ]" & _
        "};"
    GetResult "let response = " & jsonResponse & _
        "let topSellers = response.users" & _
        "  .filter(fun(u) { return u.active && u.sales > 10000 })" & _
        "  .map(fun(u) { return { name: u.name, bonus: u.sales * 0.1 } })" & _
        "  .sort(fun(a, b) {" & _
        "    if (a.bonus > b.bonus) { return -1 };" & _
        "    if (a.bonus < b.bonus) { return 1 };" & _
        "    return 0;" & _
        "  });" & _
        "print(topSellers);", True
    Set Globals = scriptEngine.GetGlobals
    With Globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.Count))
    End With
    expected = "PRINT:[ { name: 'Charlie', bonus: 2200 }, { name: 'Alice', bonus: 1500 } ]"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub injecting_asf_objects()
    On Error GoTo TestFail
    Dim Globals As ASF_Globals
    Dim engine As ASF: Set engine = New ASF
    Dim obj As ASF_Map
    With engine
        .Run .Compile( _
        "return {" & _
        "  users: [" & _
        "    { 'id': 1, 'name': 'Alice', 'sales': 15000, 'active': true }," & _
        "    { 'id': 2, 'name': 'Bob', 'sales': 8000, 'active': false }," & _
        "    { 'id': 3, 'name': 'Charlie', 'sales': 22000, 'active': true }" & _
        "  ]" & _
        "};")
        Set obj = .OUTPUT_
    End With
    Set scriptEngine = New ASF
    With scriptEngine
        .verbose = True
        .InjectVariable "response", obj
        .Run .Compile( _
        "let topSellers = response.users" & _
        "  .filter(fun(u) { return u.active && u.sales > 10000 })" & _
        "  .map(fun(u) { return { name: u.name, bonus: u.sales * 0.1 } })" & _
        "  .sort(fun(a, b) {" & _
        "    if (a.bonus > b.bonus) { return -1 };" & _
        "    if (a.bonus < b.bonus) { return 1 };" & _
        "    return 0;" & _
        "  });" & _
        "print(topSellers);")
    End With
    Set Globals = scriptEngine.GetGlobals
    With Globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.Count))
    End With
    expected = "PRINT:[ { name: 'Charlie', bonus: 2200 }, { name: 'Alice', bonus: 1500 } ]"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub
Private Sub spread_simple()
    On Error GoTo TestFail
    Dim Globals As ASF_Globals
    GetResult "arr1 = [1, 2, 3]; arr2 = [0, ...arr1, 4]; print(arr2)", True
    Set Globals = scriptEngine.GetGlobals
    With Globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.Count))
    End With
    expected = "PRINT:[ 0, 1, 2, 3, 4 ]"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub
Private Sub spread_multiple()
    On Error GoTo TestFail
    Dim Globals As ASF_Globals
    GetResult "a = [1, 2]; b = [3, 4]; c = [...a, ...b, 5]; print(c)", True
    Set Globals = scriptEngine.GetGlobals
    With Globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.Count))
    End With
    expected = "PRINT:[ 1, 2, 3, 4, 5 ]"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub
Private Sub spread_with_empty()
    On Error GoTo TestFail
    Dim Globals As ASF_Globals
    GetResult "empty = []; withEmpty = [1, ...empty, 2]; print(withEmpty)", True
    Set Globals = scriptEngine.GetGlobals
    With Globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.Count))
    End With
    expected = "PRINT:[ 1, 2 ]"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub
Private Sub spread_strings()
    On Error GoTo TestFail
    Dim Globals As ASF_Globals
    GetResult "chars = [...'hello']; print(chars)", True
    Set Globals = scriptEngine.GetGlobals
    With Globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.Count))
    End With
    expected = "PRINT:[ 'h', 'e', 'l', 'l', 'o' ]"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub
Private Sub spread_mixed_string()
    On Error GoTo TestFail
    Dim Globals As ASF_Globals
    GetResult "result = [1, ...'ab', 2]; print(result)", True
    Set Globals = scriptEngine.GetGlobals
    With Globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.Count))
    End With
    expected = "PRINT:[ 1, 'a', 'b', 2 ]"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub
Private Sub spread_combined()
    On Error GoTo TestFail
    Dim Globals As ASF_Globals
    GetResult "begin = [1]; middle = [2, 3, 4]; end = [5]; combined = [...begin, ...middle, ...end]; print(combined)", True
    Set Globals = scriptEngine.GetGlobals
    With Globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.Count))
    End With
    expected = "PRINT:[ 1, 2, 3, 4, 5 ]"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub
Private Sub spread_nested()
    On Error GoTo TestFail
    Dim Globals As ASF_Globals
    GetResult "nested = [[1, 2], [3, 4]]; spread = [...nested]; print(spread)", True
    Set Globals = scriptEngine.GetGlobals
    With Globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.Count))
    End With
    expected = "PRINT:[ [ 1, 2 ], [ 3, 4 ] ]"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub
Private Sub spread_shallow_copy()
    On Error GoTo TestFail
    Dim Globals As ASF_Globals
    GetResult "original = [1, 2, 3]; copy = [...original]; copy[1] = 999; print(original[1]); print(copy[1])", True
    Set Globals = scriptEngine.GetGlobals
    With Globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.Count - 1)) & ", " & _
                CStr(.gRuntimeLog(.gRuntimeLog.Count))
    End With
    expected = "PRINT:1, PRINT:999"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub
Private Sub spread_basic_function_call()
    On Error GoTo TestFail
    actual = CStr(GetResult("fun add(a, b, c) { return a + b + c; }; nums = [1, 2, 3]; result = add(...nums); return result"))
    expected = "6"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub
Private Sub spread_many_in_function_call()
    On Error GoTo TestFail
    actual = CStr(GetResult("fun sum5(a, b, c, d, e) { return a + b + c + d + e; };" & _
                            "arr1 = [1, 2]; arr2 = [3, 4]; total = sum5(...arr1, ...arr2, 5); return total"))
    expected = "15"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub
Private Sub spread_objects_basic()
    On Error GoTo TestFail
    actual = CStr(GetResult("obj1 = {a: 1, b: 2}; obj2 = {c: 3, ...obj1, d: 4};" & _
                            "return `${obj2.a}; ${obj2.b}; ${obj2.c}; ${obj2.d}`"))
    expected = "1; 2; 3; 4"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub
Private Sub spread_objects_override()
    On Error GoTo TestFail
    actual = CStr(GetResult("defaults = {x: 1, y: 2}; override = {...defaults, x: 10};" & _
                            "return `${override.x}; ${override.y}`"))
    expected = "10; 2"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub
Private Sub spread_multiple_objects()
    On Error GoTo TestFail
    actual = CStr(GetResult("obj1 = {a: 1}; obj2 = {a: 2, b: 3}; merged = {...obj1, ...obj2};" & _
                            "return `${merged.a}; ${merged.b}`"))
    expected = "2; 3"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub
Private Sub spread_array_in_object()
    On Error GoTo TestFail
    actual = CStr(GetResult("arr = ['x', 'y', 'z']; obj = {...arr};" & _
                            "return `${obj}`"))
    expected = "{ 1: 'x', 2: 'y', 3: 'z' }"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub
Private Sub spread_string_in_object()
    On Error GoTo TestFail
    actual = CStr(GetResult("obj = {...'ab'};" & _
                            "return `${obj}`"))
    expected = "{ 1: 'a', 2: 'b' }"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub
Private Sub spread_objects_with_primitives()
    On Error GoTo TestFail
    actual = CStr(GetResult("obj1 = {...null}; obj2 = {...undefined}; obj3 = {...42}; obj4 = {...true};" & _
                            "merged = {a: 1, ...obj1, ...obj2, ...obj3, ...obj4, b: 2}; return `${merged}`"))
    expected = "{ a: 1, b: 2 }"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub
Private Sub rest_parameter_basic()
    On Error GoTo TestFail
    actual = CStr(GetResult("fun sum(...numbers) {" & _
                            "    total = 0;" & _
                            "    for (n of numbers) {" & _
                            "        total = total + n;" & _
                            "    };" & _
                            "    return total;" & _
                            "};" & _
                            "result = sum(1, 2, 3, 4, 5);" & _
                            "return result"))
    expected = "15"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub
Private Sub rest_parameter_no_arguments()
    On Error GoTo TestFail
    actual = CStr(GetResult("fun test(...args) {" & _
                            "    return args.length();" & _
                            "};" & _
                            "result = test();" & _
                            "return result"))
    expected = "0"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub
Private Sub rest_parameter_with_regular_parameters()
    On Error GoTo TestFail
    actual = CStr(GetResult("fun greetAll(greeting, ...names) {" & _
                            "    result = greeting + ': ';" & _
                            "    for (name of names) {" & _
                            "        result = result + name + ', ';" & _
                            "    };" & _
                            "    return result.slice(0, -2);" & _
                            "};" & _
                            "msg = greetAll('Hello', 'Alice', 'Bob', 'Charlie');" & _
                            "return msg"))
    expected = "Hello: Alice, Bob, Charlie"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub
Private Sub rest_parameter_with_single_extra_argument()
    On Error GoTo TestFail
    actual = CStr(GetResult("fun test(a, b, ...rest) {" & _
                            "    return rest;" & _
                            "};" & _
                            "result = test(1, 2, 3);" & _
                            "return `${result}`"))
    expected = "[ 3 ]"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub
Private Sub rest_parameter_not_at_end()
    On Error GoTo TestFail
    actual = CStr(GetResult("fun test(a, ...rest, b) {" & _
                            "    return rest;" & _
                            "};" & _
                            "result = test(1, 2, 3);" & _
                            "return `${result}`"))
    expected = ""
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub
Private Sub rest_parameter_with_exact_number_of_arguments()
    On Error GoTo TestFail
    actual = CStr(GetResult("fun test(a, b, ...rest) {" & _
                            "    return rest.length();" & _
                            "};" & _
                            "result = test(1, 2);" & _
                            "return result"))
    expected = "0"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub
Private Sub rest_parameter_receiving_expread_arguments()
    On Error GoTo TestFail
    actual = CStr(GetResult("fun sumThree(a, b, c) {" & _
                            "    return a + b + c;" & _
                            "};" & _
                            "parts1 = [1];" & _
                            "parts2 = [2, 3];" & _
                            "fullArgs = [...parts1, ...parts2];" & _
                            "return sumThree(...fullArgs)"))
    expected = "6"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub
Private Sub rest_parameter_and_expread_chain()
    On Error GoTo TestFail
    actual = CStr(GetResult("fun combine(...arrays) {" & _
                            "    result = [];" & _
                            "    for (arr of arrays) {" & _
                            "        result = [...result, ...arr];" & _
                            "    };" & _
                            "    return result;" & _
                            "};" & _
                            "return `${combine([1, 2], [3, 4], [5, 6])}`;"))
    expected = "[ 1, 2, 3, 4, 5, 6 ]"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub
Private Sub array_destructuring_basic()
    On Error GoTo TestFail
    actual = CStr(GetResult("[a, b, c] = [1, 2, 3];" & _
                            "return `${a}; ${b}; ${c}`;"))
    expected = "1; 2; 3"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub
Private Sub array_destructuring_fewer_targets()
    On Error GoTo TestFail
    actual = CStr(GetResult("[x, y] = [10, 20, 30, 40];" & _
                            "return `${x}; ${y}`;"))
    expected = "10; 20"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub
Private Sub array_destructuring_more_targets()
    On Error GoTo TestFail
    actual = CStr(GetResult("[p, q, r] = [100, 200];" & _
                            "return typeof r;"))
    expected = "undefined"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub
Private Sub array_destructuring_with_rest()
    On Error GoTo TestFail
    actual = CStr(GetResult("[first, second, ...rest] = [1, 2, 3, 4, 5];" & _
                            "return `${rest}`;"))
    expected = "[ 3, 4, 5 ]"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub
Private Sub array_destructuring_with_no_remaining_elm()
    On Error GoTo TestFail
    actual = CStr(GetResult("[a, b, ...others] = [10, 20];" & _
                            "return others.length;"))
    expected = "0"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub
Private Sub module_system_working_directory()
    Dim wd As String
    Dim eng As New ASF
    wd = ThisWorkbook.path
    With eng
        .InjectVariable "wd", wd
        On Error GoTo TestFail
        actual = CStr(.Run(.Compile("scwd(wd); return cwd()")))
    End With
    expected = wd
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub
Private Sub module_system_named_exports_imports()
    Dim wd As String
    Dim eng As New ASF
    wd = ThisWorkbook.path
    With eng
        .InjectVariable "wd", wd
        On Error GoTo TestFail
        actual = CStr(.Execute(wd & "\main_math.vas"))
    End With
    expected = "5 + 3 = 8, Circle area: 78.53975"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub
Private Sub module_system_default_export()
    Dim wd As String
    Dim eng As New ASF
    wd = ThisWorkbook.path
    With eng
        .InjectVariable "wd", wd
        On Error GoTo TestFail
        actual = CStr(.Execute(wd & "\main_calculator.vas"))
    End With
    expected = "15"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub
Private Sub module_system_namespace_import()
    Dim wd As String
    Dim eng As New ASF
    wd = ThisWorkbook.path
    With eng
        .InjectVariable "wd", wd
        On Error GoTo TestFail
        actual = CStr(.Execute(wd & "\main_utils.vas"))
    End With
    expected = "JOHN DOE"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub
Private Sub module_system_mixed_imports()
    Dim wd As String
    Dim eng As New ASF
    wd = ThisWorkbook.path
    With eng
        .InjectVariable "wd", wd
        On Error GoTo TestFail
        actual = CStr(.Execute(wd & "\app.vas"))
    End With
    expected = "Main function | Helper function | Version: 1.0.0"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub
Private Sub module_system_prototype_imports()
    Dim wd As String
    Dim eng As New ASF
    wd = ThisWorkbook.path
    With eng
        .AppAccess = True
        .InjectVariable "wd", wd
        On Error GoTo TestFail
        actual = CLng(.Execute(wd & "\main_prototype.vas", ThisWorkbook))
        Debug.Print .GetCallStackTrace
    End With
    expected = CLng(255)
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub module_system_prototype_single_import()
    ' importing one prototype still activates the library's others (highlight calls addStyle)
    Dim wd As String
    Dim eng As New ASF
    wd = ThisWorkbook.path
    With eng
        .AppAccess = True
        .InjectVariable "wd", wd
        On Error GoTo TestFail
        actual = CLng(.Execute(wd & "\main_prototype_single.vas", ThisWorkbook))
    End With
    expected = CLng(255)
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub
Private Sub module_system_prototype_namespace_import()
    ' a namespace import activates the library's prototypes
    Dim wd As String
    Dim eng As New ASF
    wd = ThisWorkbook.path
    With eng
        .AppAccess = True
        .InjectVariable "wd", wd
        On Error GoTo TestFail
        actual = CLng(.Execute(wd & "\main_prototype_namespace.vas", ThisWorkbook))
    End With
    expected = CLng(255)
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub
Private Sub module_system_prototype_side_effect_import()
    ' a side-effect import activates the library's prototypes
    Dim wd As String
    Dim eng As New ASF
    wd = ThisWorkbook.path
    With eng
        .AppAccess = True
        .InjectVariable "wd", wd
        On Error GoTo TestFail
        actual = CLng(.Execute(wd & "\main_prototype_side_effect.vas", ThisWorkbook))
    End With
    expected = CLng(65535)
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub
Private Sub math_basic()
    On Error GoTo TestFail
    
    actual = CStr(GetResult("return(Math.pow(Math.min(1, 2, -3), 2));"))
    expected = "9"
    AssertEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub calling_native_function_app_level()
    On Error GoTo TestFail
    Set scriptEngine = New ASF
    Dim progIdx As Long
    
    With scriptEngine
        Dim col As New Collection
        col.Add 5, "myNumber"
        .AppAccess = True
        progIdx = .Compile("/*Get collection item*/ return($1.item('myNumber'))")
        .Run progIdx, col
        actual = .OUTPUT_
    End With
    expected = 5
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub calling_native_function_app_level_2()
    On Error GoTo TestFail
    Set scriptEngine = New ASF
    Dim progIdx As Long
    
    With scriptEngine
        .AppAccess = True
        progIdx = .Compile("/*Actibe Workbook Name*/ return($1.sheets.count)")
        .Run progIdx, ThisWorkbook
        actual = .OUTPUT_
    End With
    expected = ThisWorkbook.Sheets().Count
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub

Private Sub calling_array_method_from_collection()
    On Error GoTo TestFail
    Set scriptEngine = New ASF
    Dim progIdx As Long
    
    With scriptEngine
        Dim col As New Collection
        col.Add "Jhon", "Name1"
        col.Add "Marie", "Name2"
        progIdx = .Compile("return $1.filter(fun(name){ return name.startsWith('M')})[1]")
        .Run progIdx, col
        actual = .OUTPUT_
    End With
    expected = "Marie"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub
Private Sub monkey_patch_worksheets_name()
    On Error GoTo TestFail
    Set scriptEngine = New ASF
    Dim progIdx As Long
    
    With scriptEngine
        .AppAccess = True
        .OverrideCollMethods = True
        progIdx = .Compile(" prototype.COM.Worksheet Name() {" & _
                            "   return 'PATCHED';" & _
                            "};" & _
                            "return $1.Name")
        .Run progIdx, ThisWorkbook.Sheets(2)
        actual = .OUTPUT_
    End With
    expected = "PATCHED"
    AssertEqual expected, actual
TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub
Private Sub nullish_coalescing()
    On Error GoTo TestFail
    
    actual = CStr(GetResult("let foo = null ?? empty ?? 'default string'; return foo;"))
    expected = "default string"
    AssertEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    AssertFail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub
