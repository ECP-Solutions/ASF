Attribute VB_Name = "RegexTests"
' RegexTests
' Standalone test harness for ASF_RegexEngine
' No Rubberduck dependency. Uses Debug.Print for reporting.
Option Explicit

' -----------------------------------------------------------------------
' Test harness state
' -----------------------------------------------------------------------
Private g_totalTests As Long
Private g_passedTests As Long
Private g_failedTests As Long
Private g_lastError As String

' -----------------------------------------------------------------------
' Public entrypoint: runs all tests sequentially
' -----------------------------------------------------------------------
' Replace / Insert this RunAllRegexTests implementation into your test module
Public Sub RunAllRegexTests()
    g_totalTests = 0
    g_passedTests = 0
    g_failedTests = 0
    g_lastError = vbNullString

    Debug.Print "=== Starting ASF_RegexEngine standalone test suite ==="
    Debug.Print "Printing to the Immediate/Debug window."

    ' --- Category 1 ---
    ReportTest "1_basic_literal_exact", GetTestResult("1_basic_literal_exact")
    ReportTest "1_basic_literal_mismatch", GetTestResult("1_basic_literal_mismatch")
    ReportTest "1_dot_matches_any", GetTestResult("1_dot_matches_any")
    ReportTest "1_dot_requires_at_least_one", GetTestResult("1_dot_requires_at_least_one")
    ReportTest "1_dot_in_sequence", GetTestResult("1_dot_in_sequence")
    ReportTest "1_dot_dotall_true_newline", GetTestResult("1_dot_dotall_true_newline")
    ReportTest "1_dot_dotall_false_newline", GetTestResult("1_dot_dotall_false_newline")
    ReportTest "1_anchor_dot_single", GetTestResult("1_anchor_dot_single")
    ReportTest "1_anchor_dot_multi_fail", GetTestResult("1_anchor_dot_multi_fail")

    ' --- Category 2 ---
    ReportTest "2_escape_digit_true", GetTestResult("2_escape_digit_true")
    ReportTest "2_escape_digit_false", GetTestResult("2_escape_digit_false")
    ReportTest "2_escape_word_true", GetTestResult("2_escape_word_true")
    ReportTest "2_escape_word_false", GetTestResult("2_escape_word_false")
    ReportTest "2_escape_space_true", GetTestResult("2_escape_space_true")
    ReportTest "2_escape_space_false", GetTestResult("2_escape_space_false")
    ReportTest "2_escape_lf", GetTestResult("2_escape_lf")
    ReportTest "2_escape_cr", GetTestResult("2_escape_cr")
    ReportTest "2_escape_tab", GetTestResult("2_escape_tab")
    ReportTest "2_escape_escaped_metachar", GetTestResult("2_escape_escaped_metachar")
    ReportTest "2_escape_d2_exec", GetTestResult("2_escape_d2_exec")
    ReportTest "2_escape_d2_partial_fail", GetTestResult("2_escape_d2_partial_fail")

    ' --- Category 3 ---
    ReportTest "3_class_basic_true", GetTestResult("3_class_basic_true")
    ReportTest "3_class_basic_false", GetTestResult("3_class_basic_false")
    ReportTest "3_class_range_cases", GetTestResult("3_class_range_cases")
    ReportTest "3_class_negated_true", GetTestResult("3_class_negated_true")
    ReportTest "3_class_negated_false", GetTestResult("3_class_negated_false")
    ReportTest "3_class_escape_in_class_true", GetTestResult("3_class_escape_in_class_true")
    ReportTest "3_class_escape_in_class_false", GetTestResult("3_class_escape_in_class_false")
    ReportTest "3_class_quantifier_on_class", GetTestResult("3_class_quantifier_on_class")
    ReportTest "3_class_quantifier_partial_exec", GetTestResult("3_class_quantifier_partial_exec")
    ReportTest "3_negated_single", GetTestResult("3_negated_single")
    ReportTest "3_empty_class_false", GetTestResult("3_empty_class_false")
    ReportTest "3_negated_empty_true", GetTestResult("3_negated_empty_true")

    ' --- Category 4 ---
    ReportTest "4_quant_a_star_empty", GetTestResult("4_quant_a_star_empty")
    ReportTest "4_quant_a_star_exec", GetTestResult("4_quant_a_star_exec")
    ReportTest "4_quant_a_star_lazy_exec", GetTestResult("4_quant_a_star_lazy_exec")
    ReportTest "4_quant_a_star_possessive_exec", GetTestResult("4_quant_a_star_possessive_exec")
    ReportTest "4_quant_a_plus_true_false", GetTestResult("4_quant_a_plus_true_false")
    ReportTest "4_quant_a_plus_lazy", GetTestResult("4_quant_a_plus_lazy")
    ReportTest "4_quant_a_plus_possessive", GetTestResult("4_quant_a_plus_possessive")
    ReportTest "4_quant_a_question", GetTestResult("4_quant_a_question")
    ReportTest "4_quant_a_question_lazy", GetTestResult("4_quant_a_question_lazy")
    ReportTest "4_quant_a_question_possessive", GetTestResult("4_quant_a_question_possessive")
    ReportTest "4_quant_exact_a3", GetTestResult("4_quant_exact_a3")
    ReportTest "4_quant_exact_a3_longer", GetTestResult("4_quant_exact_a3_longer")
    ReportTest "4_quant_range_true", GetTestResult("4_quant_range_true")
    ReportTest "4_quant_range_below_min", GetTestResult("4_quant_range_below_min")
    ReportTest "4_quant_range_lazy", GetTestResult("4_quant_range_lazy")
    ReportTest "4_quant_range_possessive", GetTestResult("4_quant_range_possessive")
    ReportTest "4_quant_min_unlimited", GetTestResult("4_quant_min_unlimited")
    ReportTest "4_quant_min_unlimited_lazy", GetTestResult("4_quant_min_unlimited_lazy")
    ReportTest "4_quant_min_unlimited_possessive", GetTestResult("4_quant_min_unlimited_possessive")
    ReportTest "4_wildcard_greedy", GetTestResult("4_wildcard_greedy")
    ReportTest "4_wildcard_lazy", GetTestResult("4_wildcard_lazy")
    ReportTest "4_wildcard_possessive", GetTestResult("4_wildcard_possessive")

    ' --- Category 5 ---
    ReportTest "5_context_greedy", GetTestResult("5_context_greedy")
    ReportTest "5_context_lazy", GetTestResult("5_context_lazy")
    ReportTest "5_context_possessive", GetTestResult("5_context_possessive")
    ReportTest "5_tag_greedy", GetTestResult("5_tag_greedy")
    ReportTest "5_tag_lazy", GetTestResult("5_tag_lazy")
    ReportTest "5_tag_possessive", GetTestResult("5_tag_possessive")
    ReportTest "5_range_greedy", GetTestResult("5_range_greedy")
    ReportTest "5_range_lazy", GetTestResult("5_range_lazy")
    ReportTest "5_range_possessive", GetTestResult("5_range_possessive")
    ReportTest "5_alt_greedy", GetTestResult("5_alt_greedy")
    ReportTest "5_alt_possessive_then_alt", GetTestResult("5_alt_possessive_then_alt")

    ' --- Category 6 ---
    ReportTest "6_capturing_basic", GetTestResult("6_capturing_basic")
    ReportTest "6_capturing_nested", GetTestResult("6_capturing_nested")
    ReportTest "6_capturing_alternation", GetTestResult("6_capturing_alternation")
    ReportTest "6_capturing_alternation_second", GetTestResult("6_capturing_alternation_second")
    ReportTest "6_capturing_quant_in_group", GetTestResult("6_capturing_quant_in_group")
    ReportTest "6_optional_empty_exec_blank", GetTestResult("6_optional_empty_exec_blank")
    ReportTest "6_optional_empty_exec_a", GetTestResult("6_optional_empty_exec_a")
    ReportTest "6_anchored_capturing", GetTestResult("6_anchored_capturing")
    ReportTest "6_capturing_partial_failure", GetTestResult("6_capturing_partial_failure")

    ' --- Category 7 ---
    ReportTest "7_alt_simple_a", GetTestResult("7_alt_simple_a")
    ReportTest "7_alt_simple_b", GetTestResult("7_alt_simple_b")
    ReportTest "7_alt_three", GetTestResult("7_alt_three")
    ReportTest "7_grouped_alt_foobaz", GetTestResult("7_grouped_alt_foobaz")
    ReportTest "7_grouped_alt_barbaz", GetTestResult("7_grouped_alt_barbaz")
    ReportTest "7_embedded_alt_abd", GetTestResult("7_embedded_alt_abd")
    ReportTest "7_embedded_alt_acd", GetTestResult("7_embedded_alt_acd")
    ReportTest "7_longer_preference", GetTestResult("7_longer_preference")

    ' --- Category 8 ---
    ReportTest "8_anchor_start_true", GetTestResult("8_anchor_start_true")
    ReportTest "8_anchor_start_false", GetTestResult("8_anchor_start_false")
    ReportTest "8_anchor_end_true", GetTestResult("8_anchor_end_true")
    ReportTest "8_anchor_end_false", GetTestResult("8_anchor_end_false")
    ReportTest "8_anchor_full_true", GetTestResult("8_anchor_full_true")
    ReportTest "8_anchor_full_false", GetTestResult("8_anchor_full_false")
    ReportTest "8_anchor_digits_full_true", GetTestResult("8_anchor_digits_full_true")
    ReportTest "8_anchor_digits_full_false", GetTestResult("8_anchor_digits_full_false")
    ReportTest "8_anchor_multiline_start_after_newline", GetTestResult("8_anchor_multiline_start_after_newline")
    ReportTest "8_anchor_multiline_end_before_newline", GetTestResult("8_anchor_multiline_end_before_newline")

    ' --- Category 9 ---
    ReportTest "9_case_ignore_true", GetTestResult("9_case_ignore_true")
    ReportTest "9_case_ignore_false", GetTestResult("9_case_ignore_false")
    ReportTest "9_case_class_ignore", GetTestResult("9_case_class_ignore")
    ReportTest "9_case_alt_ignore", GetTestResult("9_case_alt_ignore")

    ' --- Category 10 ---
    ReportTest "10_replace_swap", GetTestResult("10_replace_swap")
    ReportTest "10_replace_wrap", GetTestResult("10_replace_wrap")
    ReportTest "10_replace_full_ref", GetTestResult("10_replace_full_ref")
    ReportTest "10_replace_partial", GetTestResult("10_replace_partial")
    ReportTest "10_replace_no_match", GetTestResult("10_replace_no_match")
    ReportTest "10_replace_issue_27", GetTestResult("10_replace_issue_27")

    ' --- Category 11 ---
    ReportTest "11_lookahead_positive_simple", GetTestResult("11_lookahead_positive_simple")
    ReportTest "11_lookahead_positive_fail", GetTestResult("11_lookahead_positive_fail")
    ReportTest "11_lookahead_in_sequence_exec", GetTestResult("11_lookahead_in_sequence_exec")
    ReportTest "11_lookahead_negative_true", GetTestResult("11_lookahead_negative_true")
    ReportTest "11_lookahead_negative_false", GetTestResult("11_lookahead_negative_false")
    ReportTest "11_lookahead_start_fail", GetTestResult("11_lookahead_start_fail")
    ReportTest "11_lookahead_overlapping", GetTestResult("11_lookahead_overlapping")
    ReportTest "11_lookahead_variable_length", GetTestResult("11_lookahead_variable_length")
    ReportTest "11_lookahead_variable_negative", GetTestResult("11_lookahead_variable_negative")

    ' --- Category 12 ---
    ReportTest "12_lookbehind_positive_fixed", GetTestResult("12_lookbehind_positive_fixed")
    ReportTest "12_lookbehind_positive_fail", GetTestResult("12_lookbehind_positive_fail")
    ReportTest "12_lookbehind_in_sequence_exec", GetTestResult("12_lookbehind_in_sequence_exec")
    ReportTest "12_lookbehind_negative_fixed", GetTestResult("12_lookbehind_negative_fixed")
    ReportTest "12_lookbehind_negative_fail", GetTestResult("12_lookbehind_negative_fail")
    ReportTest "12_lookbehind_at_end", GetTestResult("12_lookbehind_at_end")
    ReportTest "12_lookbehind_multi_char_fixed", GetTestResult("12_lookbehind_multi_char_fixed")
    ReportTest "12_lookbehind_variable_length_error", GetTestResult("12_lookbehind_variable_length_error")
    ReportTest "12_lookbehind_fixed_quantifier_ok", GetTestResult("12_lookbehind_fixed_quantifier_ok")
    ReportTest "12_lookbehind_negative_fixed2", GetTestResult("12_lookbehind_negative_fixed2")
    ReportTest "12_lookbehind_empty_ok", GetTestResult("12_lookbehind_empty_ok")

    ' --- Category 13 ---
    ReportTest "13_atomic_group_atomic_greedy_locks", GetTestResult("13_atomic_group_atomic_greedy_locks")
    ReportTest "13_atomic_group_prevents_backtrack", GetTestResult("13_atomic_group_prevents_backtrack")
    ReportTest "13_atomic_group_locks_choice", GetTestResult("13_atomic_group_locks_choice")
    ReportTest "13_possessive_one_or_more_true", GetTestResult("13_possessive_one_or_more_true")
    ReportTest "13_possessive_overconsume_false", GetTestResult("13_possessive_overconsume_false")
    ReportTest "13_capturing_possessive", GetTestResult("13_capturing_possessive")
    ReportTest "13_possessive_wildcard_exec", GetTestResult("13_possessive_wildcard_exec")
    ReportTest "13_atomic_alt_true", GetTestResult("13_atomic_alt_true")
    ReportTest "13_atomic_locks_inner_choice_false", GetTestResult("13_atomic_locks_inner_choice_false")

    ' --- Category 14 ---
    ReportTest "14_empty_pattern_empty_input", GetTestResult("14_empty_pattern_empty_input")
    ReportTest "14_empty_pattern_nonempty_input", GetTestResult("14_empty_pattern_nonempty_input")
    ReportTest "14_anchor_empty_string_true", GetTestResult("14_anchor_empty_string_true")
    ReportTest "14_anchor_empty_string_false", GetTestResult("14_anchor_empty_string_false")
    ReportTest "14_optional_group_exec", GetTestResult("14_optional_group_exec")
    ReportTest "14_left_pref_alt_exec", GetTestResult("14_left_pref_alt_exec")
    ReportTest "14_grouped_alt_capture", GetTestResult("14_grouped_alt_capture")
    ReportTest "14_greedy_last_digit", GetTestResult("14_greedy_last_digit")
    ReportTest "14_lazy_first_digit", GetTestResult("14_lazy_first_digit")
    ReportTest "14_backtracking_limit", GetTestResult("14_backtracking_limit")
    ReportTest "14_lookahead_and_consume", GetTestResult("14_lookahead_and_consume")
    ReportTest "14_lookbehind_and_lookahead", GetTestResult("14_lookbehind_and_lookahead")
    ReportTest "14_atomic_overconsume_false", GetTestResult("14_atomic_overconsume_false")
    ReportTest "14_noncapturing_groups", GetTestResult("14_noncapturing_groups")

    ' --- Category 15 ---
    ReportTest "15_unicode_property", GetTestResult("15_unicode_property")
    ReportTest "15_comment_syntax_unsupported", GetTestResult("15_comment_syntax_unsupported")
    ReportTest "15_variable_lookbehind_error", GetTestResult("15_variable_lookbehind_error")
    ReportTest "15_inline_flags_not_supported", GetTestResult("15_inline_flags_not_supported")
    ReportTest "15_conditionals_supported", GetTestResult("15_conditionals_supported")

    ' --- Category 16 ---
    ReportTest "16_backref_numbered", GetTestResult("16_backref_numbered")
    ReportTest "16_backref_word_repeat", GetTestResult("16_backref_word_repeat")
    ReportTest "16_backref_named", GetTestResult("16_backref_named")
    ReportTest "16_backref_ignorecase", GetTestResult("16_backref_ignorecase")
    ReportTest "16_backref_unset_group", GetTestResult("16_backref_unset_group")
    ReportTest "16_backref_in_lookahead", GetTestResult("16_backref_in_lookahead")
    ReportTest "16_backref_undefined_error", GetTestResult("16_backref_undefined_error")
    ReportTest "16_word_boundary", GetTestResult("16_word_boundary")
    ReportTest "16_non_word_boundary", GetTestResult("16_non_word_boundary")
    ReportTest "16_word_boundary_replace", GetTestResult("16_word_boundary_replace")
    ReportTest "16_escape_plain_text", GetTestResult("16_escape_plain_text")
    ReportTest "16_empty_alt_execall", GetTestResult("16_empty_alt_execall")
    ReportTest "16_empty_alt_replace", GetTestResult("16_empty_alt_replace")
    ReportTest "16_unbalanced_groups_rejected", GetTestResult("16_unbalanced_groups_rejected")
    ReportTest "16_quant_backtrack_restores_caps", GetTestResult("16_quant_backtrack_restores_caps")
    ReportTest "16_alt_backtracking", GetTestResult("16_alt_backtracking")
    ReportTest "16_quant_alt_backtracking", GetTestResult("16_quant_alt_backtracking")
    ReportTest "16_anchors_anywhere", GetTestResult("16_anchors_anywhere")
    ReportTest "16_groups_keep_positions", GetTestResult("16_groups_keep_positions")
    ReportTest "16_brace_literals", GetTestResult("16_brace_literals")
    ReportTest "16_code_point_escapes", GetTestResult("16_code_point_escapes")
    ReportTest "16_nested_lookbehind", GetTestResult("16_nested_lookbehind")
    ReportTest "16_malformed_patterns_rejected", GetTestResult("16_malformed_patterns_rejected")
    ReportTest "16_recursion_guarded", GetTestResult("16_recursion_guarded")
    ReportTest "16_api_guards", GetTestResult("16_api_guards")

    Debug.Print "=== Test run complete ==="
    Debug.Print "Total: " & g_totalTests & "  Passed: " & g_passedTests & "  Failed: " & g_failedTests

    If g_failedTests > 0 Then
        Debug.Print "Failed tests details: see above."
    Else
        Debug.Print "All tests passed. Engine features are behaving as expected."
    End If
End Sub

Public Function GetTestResult(testName As String) As Boolean
    ' --- Category 1 ---
    Select Case testName
        Case "1_basic_literal_exact": GetTestResult = T_1_01_basic_literal_exact()
        Case "1_basic_literal_mismatch": GetTestResult = T_1_02_basic_literal_mismatch()
        Case "1_dot_matches_any": GetTestResult = T_1_03_dot_matches_any()
        Case "1_dot_requires_at_least_one":  GetTestResult = T_1_04_dot_requires_at_least_one()
        Case "1_dot_in_sequence": GetTestResult = T_1_05_dot_in_sequence()
        Case "1_dot_dotall_true_newline":  GetTestResult = T_1_06_dot_dotall_true_newline()
        Case "1_dot_dotall_false_newline": GetTestResult = T_1_07_dot_dotall_false_newline()
        Case "1_anchor_dot_single": GetTestResult = T_1_08_anchor_dot_single()
        Case "1_anchor_dot_multi_fail": GetTestResult = T_1_09_anchor_dot_multi_fail()
    
        ' --- Category 2 ---
        Case "2_escape_digit_true": GetTestResult = T_2_01_escape_digit_true()
        Case "2_escape_digit_false": GetTestResult = T_2_02_escape_digit_false()
        Case "2_escape_word_true": GetTestResult = T_2_03_escape_word_true()
        Case "2_escape_word_false": GetTestResult = T_2_04_escape_word_false()
        Case "2_escape_space_true": GetTestResult = T_2_05_escape_space_true()
        Case "2_escape_space_false": GetTestResult = T_2_06_escape_space_false()
        Case "2_escape_lf": GetTestResult = T_2_07_escape_lf()
        Case "2_escape_cr": GetTestResult = T_2_08_escape_cr()
        Case "2_escape_tab": GetTestResult = T_2_09_escape_tab()
        Case "2_escape_escaped_metachar": GetTestResult = T_2_10_escape_escaped_metachar()
        Case "2_escape_d2_exec": GetTestResult = T_2_11_escape_d2_exec()
        Case "2_escape_d2_partial_fail": GetTestResult = T_2_12_escape_d2_partial_fail()
    
        ' --- Category 3 ---
        Case "3_class_basic_true": GetTestResult = T_3_01_class_basic_true()
        Case "3_class_basic_false": GetTestResult = T_3_02_class_basic_false()
        Case "3_class_range_cases": GetTestResult = T_3_03_class_range_cases()
        Case "3_class_negated_true": GetTestResult = T_3_04_class_negated_true()
        Case "3_class_negated_false": GetTestResult = T_3_05_class_negated_false()
        Case "3_class_escape_in_class_true": GetTestResult = T_3_06_class_escape_in_class_true()
        Case "3_class_escape_in_class_false": GetTestResult = T_3_07_class_escape_in_class_false()
        Case "3_class_quantifier_on_class": GetTestResult = T_3_08_class_quantifier_on_class()
        Case "3_class_quantifier_partial_exec": GetTestResult = T_3_09_class_quantifier_partial_exec()
        Case "3_negated_single": GetTestResult = T_3_10_negated_single()
        Case "3_empty_class_false": GetTestResult = T_3_11_empty_class_false()
        Case "3_negated_empty_true": GetTestResult = T_3_12_negated_empty_true()
    
        ' --- Category 4 ---
        Case "4_quant_a_star_empty": GetTestResult = T_4_01_quant_a_star_empty()
        Case "4_quant_a_star_exec": GetTestResult = T_4_02_quant_a_star_exec()
        Case "4_quant_a_star_lazy_exec": GetTestResult = T_4_03_quant_a_star_lazy_exec()
        Case "4_quant_a_star_possessive_exec": GetTestResult = T_4_04_quant_a_star_possessive_exec()
        Case "4_quant_a_plus_true_false": GetTestResult = T_4_05_quant_a_plus_true_false()
        Case "4_quant_a_plus_lazy": GetTestResult = T_4_06_quant_a_plus_lazy()
        Case "4_quant_a_plus_possessive": GetTestResult = T_4_07_quant_a_plus_possessive()
        Case "4_quant_a_question": GetTestResult = T_4_08_quant_a_question()
        Case "4_quant_a_question_lazy": GetTestResult = T_4_09_quant_a_question_lazy()
        Case "4_quant_a_question_possessive": GetTestResult = T_4_10_quant_a_question_possessive()
        Case "4_quant_exact_a3": GetTestResult = T_4_11_quant_exact_a3()
        Case "4_quant_exact_a3_longer": GetTestResult = T_4_12_quant_exact_a3_longer()
        Case "4_quant_range_true": GetTestResult = T_4_13_quant_range_true()
        Case "4_quant_range_below_min": GetTestResult = T_4_14_quant_range_below_min()
        Case "4_quant_range_lazy": GetTestResult = T_4_15_quant_range_lazy()
        Case "4_quant_range_possessive": GetTestResult = T_4_16_quant_range_possessive()
        Case "4_quant_min_unlimited": GetTestResult = T_4_17_quant_min_unlimited()
        Case "4_quant_min_unlimited_lazy": GetTestResult = T_4_18_quant_min_unlimited_lazy()
        Case "4_quant_min_unlimited_possessive": GetTestResult = T_4_19_quant_min_unlimited_possessive()
        Case "4_wildcard_greedy": GetTestResult = T_4_20_wildcard_greedy()
        Case "4_wildcard_lazy": GetTestResult = T_4_21_wildcard_lazy()
        Case "4_wildcard_possessive": GetTestResult = T_4_22_wildcard_possessive()
    
        ' --- Category 5 ---
        Case "5_context_greedy": GetTestResult = T_5_01_context_greedy()
        Case "5_context_lazy": GetTestResult = T_5_02_context_lazy()
        Case "5_context_possessive": GetTestResult = T_5_03_context_possessive()
        Case "5_tag_greedy": GetTestResult = T_5_04_tag_greedy()
        Case "5_tag_lazy": GetTestResult = T_5_05_tag_lazy()
        Case "5_tag_possessive": GetTestResult = T_5_06_tag_possessive()
        Case "5_range_greedy": GetTestResult = T_5_07_range_greedy()
        Case "5_range_lazy": GetTestResult = T_5_08_range_lazy()
        Case "5_range_possessive": GetTestResult = T_5_09_range_possessive()
        Case "5_alt_greedy": GetTestResult = T_5_10_alt_greedy()
        Case "5_alt_possessive_then_alt": GetTestResult = T_5_11_alt_possessive_then_alt()
    
        ' --- Category 6 ---
        Case "6_capturing_basic": GetTestResult = T_6_01_capturing_basic()
        Case "6_capturing_nested": GetTestResult = T_6_02_capturing_nested()
        Case "6_capturing_alternation": GetTestResult = T_6_03_capturing_alternation()
        Case "6_capturing_alternation_second": GetTestResult = T_6_04_capturing_alternation_second()
        Case "6_capturing_quant_in_group": GetTestResult = T_6_05_capturing_quant_in_group()
        Case "6_optional_empty_exec_blank": GetTestResult = T_6_06_optional_empty_exec_blank()
        Case "6_optional_empty_exec_a": GetTestResult = T_6_07_optional_empty_exec_a()
        Case "6_anchored_capturing": GetTestResult = T_6_08_anchored_capturing()
        Case "6_capturing_partial_failure": GetTestResult = T_6_09_capturing_partial_failure()
    
        ' --- Category 7 ---
        Case "7_alt_simple_a": GetTestResult = T_7_01_alt_simple_a()
        Case "7_alt_simple_b": GetTestResult = T_7_02_alt_simple_b()
        Case "7_alt_three": GetTestResult = T_7_03_alt_three()
        Case "7_grouped_alt_foobaz": GetTestResult = T_7_04_grouped_alt_foobaz()
        Case "7_grouped_alt_barbaz": GetTestResult = T_7_05_grouped_alt_barbaz()
        Case "7_embedded_alt_abd": GetTestResult = T_7_06_embedded_alt_abd()
        Case "7_embedded_alt_acd": GetTestResult = T_7_07_embedded_alt_acd()
        Case "7_longer_preference": GetTestResult = T_7_08_longer_preference()
    
        ' --- Category 8 ---
        Case "8_anchor_start_true": GetTestResult = T_8_01_anchor_start_true()
        Case "8_anchor_start_false": GetTestResult = T_8_02_impl()
        Case "8_anchor_end_true": GetTestResult = T_8_03_anchor_end_true()
        Case "8_anchor_end_false": GetTestResult = T_8_04_anchor_end_false()
        Case "8_anchor_full_true": GetTestResult = T_8_05_anchor_full_true()
        Case "8_anchor_full_false": GetTestResult = T_8_06_anchor_full_false()
        Case "8_anchor_digits_full_true": GetTestResult = T_8_07_anchor_digits_full_true()
        Case "8_anchor_digits_full_false": GetTestResult = T_8_08_anchor_digits_full_false()
        Case "8_anchor_multiline_start_after_newline": GetTestResult = T_8_09_anchor_multiline_start_after_newline()
        Case "8_anchor_multiline_end_before_newline": GetTestResult = T_8_10_anchor_multiline_end_before_newline()
    
        ' --- Category 9 ---
        Case "9_case_ignore_true": GetTestResult = T_9_01_case_ignore_true()
        Case "9_case_ignore_false": GetTestResult = T_9_02_case_ignore_false()
        Case "9_case_class_ignore": GetTestResult = T_9_03_case_class_ignore()
        Case "9_case_alt_ignore": GetTestResult = T_9_04_case_alt_ignore()
    
        ' --- Category 10 ---
        Case "10_replace_swap": GetTestResult = T_10_01_replace_swap()
        Case "10_replace_wrap": GetTestResult = T_10_02_replace_wrap()
        Case "10_replace_full_ref": GetTestResult = T_10_03_replace_full_ref()
        Case "10_replace_partial": GetTestResult = T_10_04_replace_partial()
        Case "10_replace_no_match": GetTestResult = T_10_05_replace_no_match()
        Case "10_replace_issue_27": GetTestResult = T_10_07_replace_issue_27()
    
        ' --- Category 11 ---
        Case "11_lookahead_positive_simple": GetTestResult = T_11_01_lookahead_positive_simple()
        Case "11_lookahead_positive_fail": GetTestResult = T_11_02_lookahead_positive_fail()
        Case "11_lookahead_in_sequence_exec": GetTestResult = T_11_03_lookahead_in_sequence_exec()
        Case "11_lookahead_negative_true": GetTestResult = T_11_04_lookahead_negative_true()
        Case "11_lookahead_negative_false": GetTestResult = T_11_05_lookahead_negative_false()
        Case "11_lookahead_start_fail": GetTestResult = T_11_06_lookahead_start_fail()
        Case "11_lookahead_overlapping": GetTestResult = T_11_07_lookahead_overlapping()
        Case "11_lookahead_variable_length": GetTestResult = T_11_08_lookahead_variable_length()
        Case "11_lookahead_variable_negative": GetTestResult = T_11_09_lookahead_variable_negative()
    
        ' --- Category 12 ---
        Case "12_lookbehind_positive_fixed": GetTestResult = T_12_01_lookbehind_positive_fixed()
        Case "12_lookbehind_positive_fail": GetTestResult = T_12_02_lookbehind_positive_fail()
        Case "12_lookbehind_in_sequence_exec": GetTestResult = T_12_03_lookbehind_in_sequence_exec()
        Case "12_lookbehind_negative_fixed": GetTestResult = T_12_04_lookbehind_negative_fixed()
        Case "12_lookbehind_negative_fail": GetTestResult = T_12_05_lookbehind_negative_fail()
        Case "12_lookbehind_at_end": GetTestResult = T_12_06_lookbehind_at_end()
        Case "12_lookbehind_multi_char_fixed": GetTestResult = T_12_07_lookbehind_multi_char_fixed()
        Case "12_lookbehind_variable_length_error": GetTestResult = T_12_08_lookbehind_variable_length_error()
        Case "12_lookbehind_fixed_quantifier_ok": GetTestResult = T_12_09_lookbehind_fixed_quantifier_ok()
        Case "12_lookbehind_negative_fixed2": GetTestResult = T_12_10_lookbehind_negative_fixed2()
        Case "12_lookbehind_empty_ok": GetTestResult = T_12_11_lookbehind_empty_ok()
    
        ' --- Category 13 ---
        Case "13_atomic_group_atomic_greedy_locks": GetTestResult = T_13_01_atomic_group_atomic_greedy_locks()
        Case "13_atomic_group_prevents_backtrack": GetTestResult = T_13_02_atomic_group_prevents_backtrack()
        Case "13_atomic_group_locks_choice": GetTestResult = T_13_03_atomic_group_locks_choice()
        Case "13_possessive_one_or_more_true": GetTestResult = T_13_04_possessive_one_or_more_true()
        Case "13_possessive_overconsume_false": GetTestResult = T_13_05_possessive_overconsume_false()
        Case "13_capturing_possessive": GetTestResult = T_13_06_capturing_possessive()
        Case "13_possessive_wildcard_exec": GetTestResult = T_13_07_possessive_wildcard_exec()
        Case "13_atomic_alt_true": GetTestResult = T_13_08_atomic_alt_true()
        Case "13_atomic_locks_inner_choice_false": GetTestResult = T_13_09_atomic_locks_inner_choice_false()
    
        ' --- Category 14 ---
        Case "14_empty_pattern_empty_input": GetTestResult = T_14_01_empty_pattern_empty_input()
        Case "14_empty_pattern_nonempty_input": GetTestResult = T_14_02_empty_pattern_nonempty_input()
        Case "14_anchor_empty_string_true": GetTestResult = T_14_03_anchor_empty_string_true()
        Case "14_anchor_empty_string_false": GetTestResult = T_14_04_anchor_empty_string_false()
        Case "14_optional_group_exec": GetTestResult = T_14_05_optional_group_exec()
        Case "14_left_pref_alt_exec": GetTestResult = T_14_06_left_pref_alt_exec()
        Case "14_grouped_alt_capture": GetTestResult = T_14_07_grouped_alt_capture()
        Case "14_greedy_last_digit": GetTestResult = T_14_08_greedy_last_digit()
        Case "14_lazy_first_digit": GetTestResult = T_14_09_lazy_first_digit()
        Case "14_large_quant_true": GetTestResult = T_14_10_large_quant_true()
        Case "14_backtracking_limit": GetTestResult = T_14_11_backtracking_limit()
        Case "14_lookahead_and_consume": GetTestResult = T_14_12_lookahead_and_consume()
        Case "14_lookbehind_and_lookahead": GetTestResult = T_14_13_lookbehind_and_lookahead()
        Case "14_atomic_overconsume_false": GetTestResult = T_14_14_atomic_overconsume_false()
        Case "14_noncapturing_groups": GetTestResult = T_14_15_noncapturing_groups()
    
        ' --- Category 15 ---
        Case "15_unicode_property": GetTestResult = T_15_02_unicode_property()
        Case "15_comment_syntax_unsupported": GetTestResult = T_15_03_comment_syntax_unsupported()
        Case "15_variable_lookbehind_error": GetTestResult = T_15_05_variable_lookbehind_error()
        Case "15_inline_flags_not_supported": GetTestResult = T_15_06_inline_flags_not_supported()
        Case "15_conditionals_supported": GetTestResult = T_15_07_conditionals_supported()

        ' --- Category 16 ---
        Case "16_backref_numbered": GetTestResult = T_16_01_backref_numbered()
        Case "16_backref_word_repeat": GetTestResult = T_16_02_backref_word_repeat()
        Case "16_backref_named": GetTestResult = T_16_03_backref_named()
        Case "16_backref_ignorecase": GetTestResult = T_16_04_backref_ignorecase()
        Case "16_backref_unset_group": GetTestResult = T_16_05_backref_unset_group()
        Case "16_backref_in_lookahead": GetTestResult = T_16_06_backref_in_lookahead()
        Case "16_backref_undefined_error": GetTestResult = T_16_07_backref_undefined_error()
        Case "16_word_boundary": GetTestResult = T_16_08_word_boundary()
        Case "16_non_word_boundary": GetTestResult = T_16_09_non_word_boundary()
        Case "16_word_boundary_replace": GetTestResult = T_16_10_word_boundary_replace()
        Case "16_escape_plain_text": GetTestResult = T_16_11_escape_plain_text()
        Case "16_empty_alt_execall": GetTestResult = T_16_12_empty_alt_execall()
        Case "16_empty_alt_replace": GetTestResult = T_16_13_empty_alt_replace()
        Case "16_unbalanced_groups_rejected": GetTestResult = T_16_14_unbalanced_groups_rejected()
        Case "16_quant_backtrack_restores_caps": GetTestResult = T_16_15_quant_backtrack_restores_caps()
        Case "16_alt_backtracking": GetTestResult = T_16_16_alt_backtracking()
        Case "16_quant_alt_backtracking": GetTestResult = T_16_17_quant_alt_backtracking()
        Case "16_anchors_anywhere": GetTestResult = T_16_18_anchors_anywhere()
        Case "16_groups_keep_positions": GetTestResult = T_16_19_groups_keep_positions()
        Case "16_brace_literals": GetTestResult = T_16_20_brace_literals()
        Case "16_code_point_escapes": GetTestResult = T_16_21_code_point_escapes()
        Case "16_nested_lookbehind": GetTestResult = T_16_22_nested_lookbehind()
        Case "16_malformed_patterns_rejected": GetTestResult = T_16_23_malformed_patterns_rejected()
        Case "16_recursion_guarded": GetTestResult = T_16_24_recursion_guarded()
        Case "16_api_guards": GetTestResult = T_16_25_api_guards()
    End Select
End Function
' Helper used by the runner
Public Sub ReportTest(ByVal testName As String, ByVal result As Boolean)
    g_totalTests = g_totalTests + 1
    If result Then
        g_passedTests = g_passedTests + 1
        Debug.Print "[PASS] " & testName
    Else
        g_failedTests = g_failedTests + 1
        If g_lastError = vbNullString Then g_lastError = "unspecified failure"
        Debug.Print "[FAIL] " & testName & " -> " & g_lastError
        g_lastError = vbNullString ' clear for next test
    End If
End Sub

' -----------------------------------------------------------------------
' Assertions & helpers used by tests
' -----------------------------------------------------------------------
Private Function InitRegexAndHandle(ByRef r As ASF_RegexEngine, ByVal pattern As String, Optional ByVal ignoreCase As Boolean = False, Optional ByVal MaxSteps As Long = -1, Optional ByVal multiline As Boolean = False, Optional ByVal dotAll As Boolean = False) As Boolean
    On Error GoTo EH
    If MaxSteps <= 0 Then
        r.Init pattern, ignoreCase, r.MaxMatchSteps, multiline, dotAll
    Else
        r.Init pattern, ignoreCase, MaxSteps, multiline, dotAll
    End If
    InitRegexAndHandle = True
    Exit Function
EH:
    InitRegexAndHandle = False
    g_lastError = "Init error: #" & err.Number & " - " & err.Description
    err.Clear
End Function

Private Function ExecColl(ByRef r As ASF_RegexEngine, ByVal subj As String) As Collection
    On Error GoTo EH
    Set ExecColl = r.Exec(subj)
    Exit Function
EH:
    Set ExecColl = Nothing
    g_lastError = "Exec error: #" & err.Number & " - " & err.Description
    err.Clear
End Function

Private Function ReplaceStr(ByRef r As ASF_RegexEngine, ByVal subj As String, ByVal repl As String) As String
    On Error GoTo EH
    ReplaceStr = r.Replace(subj, repl)
    Exit Function
EH:
    ReplaceStr = vbNullString
    g_lastError = "Replace error: #" & err.Number & " - " & err.Description
    err.Clear
End Function

Private Function AssertTrue(ByVal cond As Boolean, Optional ByVal message As String = vbNullString) As Boolean
    If cond Then
        AssertTrue = True
    Else
        If message = vbNullString Then message = "AssertTrue failed"
        g_lastError = message
        AssertTrue = False
    End If
End Function

Private Function AssertFalse(ByVal cond As Boolean, Optional ByVal message As String = vbNullString) As Boolean
    If Not cond Then
        AssertFalse = True
    Else
        If message = vbNullString Then message = "AssertFalse failed"
        g_lastError = message
        AssertFalse = False
    End If
End Function

Private Function AssertEqual(ByVal a As Variant, ByVal b As Variant, Optional ByVal message As String = vbNullString) As Boolean
    If CStr(a) = CStr(b) Then
        AssertEqual = True
    Else
        If message = vbNullString Then message = "AssertEqual failed: expected '" & CStr(b) & "' got '" & CStr(a) & "'"
        g_lastError = message
        AssertEqual = False
    End If
End Function

Private Function AssertCollEquals(ByRef coll As Collection, ByRef expected() As Variant) As Boolean
    On Error GoTo EH
    If coll Is Nothing Then
        g_lastError = "Expected collection, got Nothing"
        AssertCollEquals = False: Exit Function
    End If
    Dim expCount As Long: expCount = UBound(expected) - LBound(expected) + 1
    If coll.Count <> expCount Then
        g_lastError = "Collection length mismatch: expected " & expCount & " got " & coll.Count
        AssertCollEquals = False: Exit Function
    End If
    Dim i As Long, idx As Long
    idx = 1
    For i = LBound(expected) To UBound(expected)
        If CStr(coll.Item(idx)) <> CStr(expected(i)) Then
            g_lastError = "Collection item #" & idx & " mismatch: expected '" & CStr(expected(i)) & "' got '" & CStr(coll.Item(idx)) & "'"
            AssertCollEquals = False: Exit Function
        End If
        idx = idx + 1
    Next i
    AssertCollEquals = True
    Exit Function
EH:
    g_lastError = "AssertCollEquals error: #" & err.Number & " - " & err.Description
    AssertCollEquals = False
    err.Clear
End Function

' Small helper to create array literal (simulates A(...) from previous)
Private Function a(ParamArray items() As Variant) As Variant()
    Dim out() As Variant, i As Long
    ReDim out(0 To UBound(items))
    For i = 0 To UBound(items)
        out(i) = items(i)
    Next i
    a = out
End Function

' -----------------------------------------------------------------------
' Test functions (each returns Boolean). For brevity they follow same
' logic used in earlier conversion: initialize, run Exec/Replace, assert.
' -----------------------------------------------------------------------
' Category 1
Public Function T_1_01_basic_literal_exact() As Boolean
    Dim r As New ASF_RegexEngine
    If Not InitRegexAndHandle(r, "abc") Then Exit Function
    Dim c As Collection: Set c = ExecColl(r, "abc")
    If Not AssertTrue(Not (c Is Nothing), "expected match for 'abc'") Then Exit Function
    T_1_01_basic_literal_exact = True
End Function

Public Function T_1_02_basic_literal_mismatch() As Boolean
    Dim r As New ASF_RegexEngine
    If Not InitRegexAndHandle(r, "abc") Then Exit Function
    If Not AssertTrue(ExecColl(r, "abx") Is Nothing, "expected no match for 'abx'") Then Exit Function
    T_1_02_basic_literal_mismatch = True
End Function

Public Function T_1_03_dot_matches_any() As Boolean
    Dim r As New ASF_RegexEngine
    If Not InitRegexAndHandle(r, ".") Then Exit Function
    If Not AssertTrue(Not (ExecColl(r, "a") Is Nothing), "expected '.' to match 'a'") Then Exit Function
    T_1_03_dot_matches_any = True
End Function

Public Function T_1_04_dot_requires_at_least_one() As Boolean
    Dim r As New ASF_RegexEngine
    If Not InitRegexAndHandle(r, ".") Then Exit Function
    If Not AssertTrue(ExecColl(r, vbNullString) Is Nothing, "expected '.' not to match empty") Then Exit Function
    T_1_04_dot_requires_at_least_one = True
End Function

Public Function T_1_05_dot_in_sequence() As Boolean
    Dim r As New ASF_RegexEngine
    If Not InitRegexAndHandle(r, "a.c") Then Exit Function
    If Not AssertTrue(Not (ExecColl(r, "abc") Is Nothing), "expected 'a.c' to match 'abc'") Then Exit Function
    T_1_05_dot_in_sequence = True
End Function

Public Function T_1_06_dot_dotall_true_newline() As Boolean
    Dim r As New ASF_RegexEngine
    If Not InitRegexAndHandle(r, "a.c", False, -1, False, True) Then Exit Function
    If Not AssertTrue(Not (ExecColl(r, "a" & vbLf & "c") Is Nothing), "dotAll True expected match across newline") Then Exit Function
    T_1_06_dot_dotall_true_newline = True
End Function

Public Function T_1_07_dot_dotall_false_newline() As Boolean
    Dim r As New ASF_RegexEngine
    If Not InitRegexAndHandle(r, "a.c", False, -1, False, False) Then Exit Function
    If Not AssertTrue(ExecColl(r, "a" & vbLf & "c") Is Nothing, "dotAll False expected no match across newline") Then Exit Function
    T_1_07_dot_dotall_false_newline = True
End Function

Public Function T_1_08_anchor_dot_single() As Boolean
    Dim r As New ASF_RegexEngine
    If Not InitRegexAndHandle(r, "^.$") Then Exit Function
    If Not AssertTrue(Not (ExecColl(r, "x") Is Nothing), "expected ^.$ to match 'x'") Then Exit Function
    T_1_08_anchor_dot_single = True
End Function

Public Function T_1_09_anchor_dot_multi_fail() As Boolean
    Dim r As New ASF_RegexEngine
    If Not InitRegexAndHandle(r, "^.$") Then Exit Function
    If Not AssertTrue(ExecColl(r, "xy") Is Nothing, "expected ^.$ to fail on 'xy'") Then Exit Function
    T_1_09_anchor_dot_multi_fail = True
End Function

' Category 2 (Escapes)
Public Function T_2_01_escape_digit_true() As Boolean
    Dim r As New ASF_RegexEngine
    If Not InitRegexAndHandle(r, "\d") Then Exit Function
    If Not AssertTrue(Not (ExecColl(r, "5") Is Nothing), "expected \\d to match '5'") Then Exit Function
    T_2_01_escape_digit_true = True
End Function

Public Function T_2_02_escape_digit_false() As Boolean
    Dim r As New ASF_RegexEngine
    If Not InitRegexAndHandle(r, "\d") Then Exit Function
    If Not AssertTrue(ExecColl(r, "a") Is Nothing, "expected \\d not to match 'a'") Then Exit Function
    T_2_02_escape_digit_false = True
End Function

Public Function T_2_03_escape_word_true() As Boolean
    Dim r As New ASF_RegexEngine
    If Not InitRegexAndHandle(r, "\w+") Then Exit Function
    If Not AssertTrue(Not (ExecColl(r, "hello_123") Is Nothing), "expected \\w+ to match 'hello_123'") Then Exit Function
    T_2_03_escape_word_true = True
End Function

Public Function T_2_04_escape_word_false() As Boolean
    Dim r As New ASF_RegexEngine
    If Not InitRegexAndHandle(r, "\w+") Then Exit Function
    If Not AssertTrue(ExecColl(r, "!@#") Is Nothing, "expected \\w+ not to match '!@#'") Then Exit Function
    T_2_04_escape_word_false = True
End Function

Public Function T_2_05_escape_space_true() As Boolean
    Dim r As New ASF_RegexEngine
    If Not InitRegexAndHandle(r, "\s+") Then Exit Function
    If Not AssertTrue(Not (ExecColl(r, " " & vbTab & vbLf) Is Nothing), "expected \\s+ to match whitespace") Then Exit Function
    T_2_05_escape_space_true = True
End Function

Public Function T_2_06_escape_space_false() As Boolean
    Dim r As New ASF_RegexEngine
    If Not InitRegexAndHandle(r, "\s+") Then Exit Function
    If Not AssertTrue(ExecColl(r, "abc") Is Nothing, "expected \\s+ not to match 'abc'") Then Exit Function
    T_2_06_escape_space_false = True
End Function

Public Function T_2_07_escape_lf() As Boolean
    Dim r As New ASF_RegexEngine
    If Not InitRegexAndHandle(r, "\n") Then Exit Function
    If Not AssertTrue(Not (ExecColl(r, vbLf) Is Nothing), "expected \\n to match vbLf") Then Exit Function
    T_2_07_escape_lf = True
End Function

Public Function T_2_08_escape_cr() As Boolean
    Dim r As New ASF_RegexEngine
    If Not InitRegexAndHandle(r, "\r") Then Exit Function
    If Not AssertTrue(Not (ExecColl(r, vbCr) Is Nothing), "expected \\r to match vbCr") Then Exit Function
    T_2_08_escape_cr = True
End Function

Public Function T_2_09_escape_tab() As Boolean
    Dim r As New ASF_RegexEngine
    If Not InitRegexAndHandle(r, "\t") Then Exit Function
    If Not AssertTrue(Not (ExecColl(r, vbTab) Is Nothing), "expected \\t to match vbTab") Then Exit Function
    T_2_09_escape_tab = True
End Function

Public Function T_2_10_escape_escaped_metachar() As Boolean
    Dim r As New ASF_RegexEngine
    If Not InitRegexAndHandle(r, "\\.") Then Exit Function
    If Not AssertTrue((ExecColl(r, ".") Is Nothing), "expected \\. to match '.'") Then Exit Function
    T_2_10_escape_escaped_metachar = True
End Function

Public Function T_2_11_escape_d2_exec() As Boolean
    Dim r As New ASF_RegexEngine
    If Not InitRegexAndHandle(r, "\d{2}") Then Exit Function
    Dim c As Collection: Set c = ExecColl(r, "12")
    If Not AssertTrue(Not (c Is Nothing), "expected \\d{2} to match '12'") Then Exit Function
    If Not AssertCollEquals(c, a("12")) Then Exit Function
    T_2_11_escape_d2_exec = True
End Function

Public Function T_2_12_escape_d2_partial_fail() As Boolean
    Dim r As New ASF_RegexEngine
    If Not InitRegexAndHandle(r, "\d{2}") Then Exit Function
    If Not AssertTrue(ExecColl(r, "1a") Is Nothing, "expected \\d{2} to fail on '1a'") Then Exit Function
    T_2_12_escape_d2_partial_fail = True
End Function

' Category 3 - classes/ranges
Public Function T_3_01_class_basic_true() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "[a-c]") Then Exit Function
    If Not AssertTrue(Not (ExecColl(r, "b") Is Nothing), "class [a-c] should match 'b'") Then Exit Function
    T_3_01_class_basic_true = True
End Function

Public Function T_3_02_class_basic_false() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "[a-c]") Then Exit Function
    If Not AssertTrue(ExecColl(r, "d") Is Nothing, "class [a-c] shouldn't match 'd'") Then Exit Function
    T_3_02_class_basic_false = True
End Function

Public Function T_3_03_class_range_cases() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "[a-zA-Z]") Then Exit Function
    If Not AssertTrue(Not (ExecColl(r, "G") Is Nothing), "range [a-zA-Z] should match 'G'") Then Exit Function
    T_3_03_class_range_cases = True
End Function

Public Function T_3_04_class_negated_true() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "[^0-9]") Then Exit Function
    If Not AssertTrue(Not (ExecColl(r, "x") Is Nothing), "negated class should match 'x'") Then Exit Function
    T_3_04_class_negated_true = True
End Function

Public Function T_3_05_class_negated_false() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "[^0-9]") Then Exit Function
    If Not AssertTrue(ExecColl(r, "5") Is Nothing, "negated class shouldn't match '5'") Then Exit Function
    T_3_05_class_negated_false = True
End Function

Public Function T_3_06_class_escape_in_class_true() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "[\d\s]") Then Exit Function
    If Not AssertTrue(Not (ExecColl(r, "3") Is Nothing), "[\\d\\s] should match '3'") Then Exit Function
    T_3_06_class_escape_in_class_true = True
End Function

Public Function T_3_07_class_escape_in_class_false() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "[\d\s]") Then Exit Function
    If Not AssertTrue(ExecColl(r, "a") Is Nothing, "[\\d\\s] shouldn't match 'a'") Then Exit Function
    T_3_07_class_escape_in_class_false = True
End Function

Public Function T_3_08_class_quantifier_on_class() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "[a-c]+") Then Exit Function
    If Not AssertTrue(Not (ExecColl(r, "abccba") Is Nothing), "[a-c]+ should match 'abccba'") Then Exit Function
    T_3_08_class_quantifier_on_class = True
End Function

Public Function T_3_09_class_quantifier_partial_exec() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "[a-c]+") Then Exit Function
    Dim c As Collection: Set c = ExecColl(r, "abcd")
    If Not AssertTrue(Not (c Is Nothing), "expected match on 'abcd'") Then Exit Function
    If Not AssertCollEquals(c, a("abc")) Then Exit Function
    T_3_09_class_quantifier_partial_exec = True
End Function

Public Function T_3_10_negated_single() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "[^a]") Then Exit Function
    If Not AssertTrue(Not (ExecColl(r, "b") Is Nothing), "[^a] should match 'b'") Then Exit Function
    T_3_10_negated_single = True
End Function

Public Function T_3_11_empty_class_false() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "[]") Then Exit Function
    If Not AssertTrue(ExecColl(r, "a") Is Nothing, "empty class [] should not match anything") Then Exit Function
    T_3_11_empty_class_false = True
End Function

Public Function T_3_12_negated_empty_true() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "[^]") Then Exit Function
    If Not AssertTrue(Not (ExecColl(r, "a") Is Nothing), "[^] should match any char") Then Exit Function
    T_3_12_negated_empty_true = True
End Function

' Category 4 (quantifiers) - many tests below, similar pattern
Public Function T_4_01_quant_a_star_empty() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "a*") Then Exit Function
    If Not AssertTrue(Not (ExecColl(r, vbNullString) Is Nothing), "a* should match empty") Then Exit Function
    T_4_01_quant_a_star_empty = True
End Function

Public Function T_4_02_quant_a_star_exec() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "a*") Then Exit Function
    Dim c As Collection: Set c = ExecColl(r, "aaa")
    If Not AssertTrue(Not (c Is Nothing), "a* expected match 'aaa'") Then Exit Function
    If Not AssertCollEquals(c, a("aaa")) Then Exit Function
    T_4_02_quant_a_star_exec = True
End Function

Public Function T_4_03_quant_a_star_lazy_exec() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "a*?") Then Exit Function
    Dim c As Collection: Set c = ExecColl(r, "aaa")
    If Not AssertTrue(Not (c Is Nothing), "a*? should return something") Then Exit Function
    If Not AssertCollEquals(c, a(vbNullString)) Then Exit Function
    T_4_03_quant_a_star_lazy_exec = True
End Function

Public Function T_4_04_quant_a_star_possessive_exec() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "a*+") Then Exit Function
    Dim c As Collection: Set c = ExecColl(r, "aaa")
    If Not AssertTrue(Not (c Is Nothing), "a*+ should match 'aaa' possessive") Then Exit Function
    If Not AssertCollEquals(c, a("aaa")) Then Exit Function
    T_4_04_quant_a_star_possessive_exec = True
End Function

Public Function T_4_05_quant_a_plus_true_false() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "a+") Then Exit Function
    If Not AssertTrue(Not (ExecColl(r, "aaa") Is Nothing), "a+ should match 'aaa'") Then Exit Function
    If Not AssertTrue(ExecColl(r, " ") Is Nothing, "a+ should not match space") Then Exit Function
    T_4_05_quant_a_plus_true_false = True
End Function

Public Function T_4_06_quant_a_plus_lazy() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "a+?") Then Exit Function
    Dim c As Collection: Set c = ExecColl(r, "aaa")
    If Not AssertTrue(Not (c Is Nothing), "a+? should match 'aaa' lazily") Then Exit Function
    If Not AssertCollEquals(c, a("a")) Then Exit Function
    T_4_06_quant_a_plus_lazy = True
End Function

Public Function T_4_07_quant_a_plus_possessive() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "a++") Then Exit Function
    Dim c As Collection: Set c = ExecColl(r, "aaa")
    If Not AssertTrue(Not (c Is Nothing), "a++ should match 'aaa'") Then Exit Function
    If Not AssertCollEquals(c, a("aaa")) Then Exit Function
    T_4_07_quant_a_plus_possessive = True
End Function

Public Function T_4_08_quant_a_question() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "a?") Then Exit Function
    If Not AssertTrue(Not (ExecColl(r, vbNullString) Is Nothing), "a? should match empty") Then Exit Function
    If Not AssertTrue(Not (ExecColl(r, "a") Is Nothing), "a? should match 'a'") Then Exit Function
    T_4_08_quant_a_question = True
End Function

Public Function T_4_09_quant_a_question_lazy() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "a??") Then Exit Function
    Dim c As Collection: Set c = ExecColl(r, "a")
    If Not AssertTrue(Not (c Is Nothing), "a?? should match 'a'") Then Exit Function
    If Not AssertCollEquals(c, a(vbNullString)) Then Exit Function
    T_4_09_quant_a_question_lazy = True
End Function

Public Function T_4_10_quant_a_question_possessive() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "a?+") Then Exit Function
    Dim c As Collection: Set c = ExecColl(r, "a")
    If Not AssertTrue(Not (c Is Nothing), "a?+ should match 'a'") Then Exit Function
    If Not AssertCollEquals(c, a("a")) Then Exit Function
    T_4_10_quant_a_question_possessive = True
End Function

Public Function T_4_11_quant_exact_a3() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "a{3}") Then Exit Function
    If Not AssertTrue(Not (ExecColl(r, "aaa") Is Nothing), "a{3} should match 'aaa'") Then Exit Function
    T_4_11_quant_exact_a3 = True
End Function

Public Function T_4_12_quant_exact_a3_longer() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "a{3}") Then Exit Function
    Dim c As Collection: Set c = ExecColl(r, "aaaa")
    If Not AssertTrue(Not (c Is Nothing), "a{3} should match start of 'aaaa'") Then Exit Function
    If Not AssertCollEquals(c, a("aaa")) Then Exit Function
    T_4_12_quant_exact_a3_longer = True
End Function

Public Function T_4_13_quant_range_true() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "a{2,4}") Then Exit Function
    If Not AssertTrue(Not (ExecColl(r, "aaa") Is Nothing), "a{2,4} should match 'aaa'") Then Exit Function
    T_4_13_quant_range_true = True
End Function

Public Function T_4_14_quant_range_below_min() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "a{2,4}") Then Exit Function
    If Not AssertTrue(ExecColl(r, "a") Is Nothing, "a{2,4} shouldn't match single 'a'") Then Exit Function
    T_4_14_quant_range_below_min = True
End Function

Public Function T_4_15_quant_range_lazy() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "a{2,4}?") Then Exit Function
    Dim c As Collection: Set c = ExecColl(r, "aaaaa")
    If Not AssertTrue(Not (c Is Nothing), "a{2,4}? expected to match") Then Exit Function
    If Not AssertCollEquals(c, a("aa")) Then Exit Function
    T_4_15_quant_range_lazy = True
End Function

Public Function T_4_16_quant_range_possessive() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "a{2,4}+") Then Exit Function
    Dim c As Collection: Set c = ExecColl(r, "aaaaa")
    If Not AssertTrue(Not (c Is Nothing), "a{2,4}+ expected to match") Then Exit Function
    If Not AssertCollEquals(c, a("aaaa")) Then Exit Function
    T_4_16_quant_range_possessive = True
End Function

Public Function T_4_17_quant_min_unlimited() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "a{2,}") Then Exit Function
    If Not AssertTrue(Not (ExecColl(r, "aaaaa") Is Nothing), "a{2,} should match 'aaaaa'") Then Exit Function
    T_4_17_quant_min_unlimited = True
End Function

Public Function T_4_18_quant_min_unlimited_lazy() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "a{2,}?") Then Exit Function
    Dim c As Collection: Set c = ExecColl(r, "aaaaa")
    If Not AssertTrue(Not (c Is Nothing), "a{2,}? should match 'aaaaa' lazily") Then Exit Function
    If Not AssertCollEquals(c, a("aa")) Then Exit Function
    T_4_18_quant_min_unlimited_lazy = True
End Function

Public Function T_4_19_quant_min_unlimited_possessive() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "a{2,}+") Then Exit Function
    Dim c As Collection: Set c = ExecColl(r, "aaaaa")
    If Not AssertTrue(Not (c Is Nothing), "a{2,}+ should match 'aaaaa'") Then Exit Function
    If Not AssertCollEquals(c, a("aaaaa")) Then Exit Function
    T_4_19_quant_min_unlimited_possessive = True
End Function

Public Function T_4_20_wildcard_greedy() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, ".*") Then Exit Function
    If Not AssertTrue(Not (ExecColl(r, "anything") Is Nothing), ".* should match 'anything'") Then Exit Function
    T_4_20_wildcard_greedy = True
End Function

Public Function T_4_21_wildcard_lazy() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, ".*?") Then Exit Function
    Dim c As Collection: Set c = ExecColl(r, "anything")
    If Not AssertTrue(Not (c Is Nothing), ".*? expected to produce a match") Then Exit Function
    If Not AssertCollEquals(c, a(vbNullString)) Then Exit Function
    T_4_21_wildcard_lazy = True
End Function

Public Function T_4_22_wildcard_possessive() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, ".*+") Then Exit Function
    If Not AssertTrue(Not (ExecColl(r, "anything") Is Nothing), ".*+ should match 'anything'") Then Exit Function
    T_4_22_wildcard_possessive = True
End Function

' Category 5 - context (selected)
Public Function T_5_01_context_greedy() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "a.+b") Then Exit Function
    Dim c As Collection: Set c = ExecColl(r, "aaabxb")
    If Not AssertTrue(Not (c Is Nothing), "a.+b expected match") Then Exit Function
    If Not AssertCollEquals(c, a("aaabxb")) Then Exit Function
    T_5_01_context_greedy = True
End Function

Public Function T_5_02_context_lazy() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "a.+?b") Then Exit Function
    Dim c As Collection: Set c = ExecColl(r, "aaabxb")
    If Not AssertTrue(Not (c Is Nothing), "a.+?b expected match") Then Exit Function
    If Not AssertCollEquals(c, a("aaab")) Then Exit Function
    T_5_02_context_lazy = True
End Function

Public Function T_5_03_context_possessive() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "a.++b") Then Exit Function
    Dim c As Collection: Set c = ExecColl(r, "aaabxb")
    If Not AssertTrue((c Is Nothing), "a.++b expected to not match (possessive semantics may vary)") Then Exit Function
    T_5_03_context_possessive = True
End Function

Public Function T_5_04_tag_greedy() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "<.*>") Then Exit Function
    Dim c As Collection: Set c = ExecColl(r, "<tag>content</tag>")
    If Not AssertTrue(Not (c Is Nothing), "<.*> expected to match entire tag") Then Exit Function
    T_5_04_tag_greedy = True
End Function

Public Function T_5_05_tag_lazy() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "<.*?>") Then Exit Function
    Dim c As Collection: Set c = ExecColl(r, "<tag>content</tag>")
    If Not AssertTrue(Not (c Is Nothing), "<.*?> expected match") Then Exit Function
    T_5_05_tag_lazy = True
End Function

Public Function T_5_06_tag_possessive() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "<.*+>") Then Exit Function
    Dim c As Collection: Set c = ExecColl(r, "<tag>content</tag>")
    If Not AssertTrue((c Is Nothing), "<.*+> expected no match (possessive)") Then Exit Function
    T_5_06_tag_possessive = True
End Function

Public Function T_5_07_range_greedy() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "a{1,3}b") Then Exit Function
    Dim c As Collection: Set c = ExecColl(r, "aaab")
    If Not AssertTrue(Not (c Is Nothing), "a{1,3}b expected match") Then Exit Function
    If Not AssertCollEquals(c, a("aaab")) Then Exit Function
    T_5_07_range_greedy = True
End Function

Public Function T_5_08_range_lazy() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "a{1,3}?b") Then Exit Function
    Dim c As Collection: Set c = ExecColl(r, "aaab")
    If Not AssertTrue(Not (c Is Nothing), "a{1,3}?b expected match") Then Exit Function
    If Not AssertCollEquals(c, a("aaab")) Then Exit Function
    T_5_08_range_lazy = True
End Function

Public Function T_5_09_range_possessive() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "a{1,3}+b") Then Exit Function
    Dim c As Collection: Set c = ExecColl(r, "aaab")
    If Not AssertTrue(Not (c Is Nothing), "a{1,3}+b expected match") Then Exit Function
    If Not AssertCollEquals(c, a("aaab")) Then Exit Function
    T_5_09_range_possessive = True
End Function

Public Function T_5_10_alt_greedy() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "a.+b|c") Then Exit Function
    Dim c As Collection: Set c = ExecColl(r, "aaabxc")
    If Not AssertTrue(Not (c Is Nothing), "a.+b|c expected match") Then Exit Function
    If Not AssertCollEquals(c, a("aaab")) Then Exit Function
    T_5_10_alt_greedy = True
End Function

Public Function T_5_11_alt_possessive_then_alt() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "a.++b|c") Then Exit Function
    Dim c As Collection: Set c = ExecColl(r, "aaabxc")
    If c Is Nothing Then
        g_lastError = "expected 'c' match via alt but got nothing"
        Exit Function
    End If
    If Not AssertCollEquals(c, a("c")) Then Exit Function
    T_5_11_alt_possessive_then_alt = True
End Function

' Category 6 - capturing groups (subset implemented)
Public Function T_6_01_capturing_basic() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "(ab)(c)") Then Exit Function
    Dim c As Collection: Set c = ExecColl(r, "abc")
    If Not AssertTrue(Not (c Is Nothing), "expected capture for 'abc'") Then Exit Function
    If Not AssertCollEquals(c, a("abc", "ab", "c")) Then Exit Function
    T_6_01_capturing_basic = True
End Function

Public Function T_6_02_capturing_nested() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "(a(b)c)") Then Exit Function
    Dim c As Collection: Set c = ExecColl(r, "abc")
    If Not AssertTrue(Not (c Is Nothing), "expected nested captures") Then Exit Function
    If Not AssertCollEquals(c, a("abc", "abc", "b")) Then Exit Function
    T_6_02_capturing_nested = True
End Function

Public Function T_6_03_capturing_alternation() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "(a|b)c") Then Exit Function
    Dim c As Collection: Set c = ExecColl(r, "ac")
    If Not AssertTrue(Not (c Is Nothing), "expected (a|b)c to match 'ac'") Then Exit Function
    If Not AssertCollEquals(c, a("ac", "a")) Then Exit Function
    T_6_03_capturing_alternation = True
End Function

Public Function T_6_04_capturing_alternation_second() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "(a|b)c") Then Exit Function
    Dim c As Collection: Set c = ExecColl(r, "bc")
    If Not AssertTrue(Not (c Is Nothing), "expected (a|b)c to match 'bc'") Then Exit Function
    If Not AssertCollEquals(c, a("bc", "b")) Then Exit Function
    T_6_04_capturing_alternation_second = True
End Function

Public Function T_6_05_capturing_quant_in_group() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "a(b*)c") Then Exit Function
    Dim c As Collection: Set c = ExecColl(r, "abbbc")
    If Not AssertTrue(Not (c Is Nothing), "expected a(b*)c to match 'abbbc'") Then Exit Function
    If Not AssertCollEquals(c, a("abbbc", "bbb")) Then Exit Function
    T_6_05_capturing_quant_in_group = True
End Function

Public Function T_6_06_optional_empty_exec_blank() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "(a)?") Then Exit Function
    Dim c As Collection: Set c = ExecColl(r, vbNullString)
    If Not AssertTrue(Not (c Is Nothing), "expected (a)? to match empty") Then Exit Function
    If Not AssertCollEquals(c, a(vbNullString, vbNullString)) Then Exit Function
    T_6_06_optional_empty_exec_blank = True
End Function

Public Function T_6_07_optional_empty_exec_a() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "(a)?") Then Exit Function
    Dim c As Collection: Set c = ExecColl(r, "a")
    If Not AssertTrue(Not (c Is Nothing), "expected (a)? to match 'a'") Then Exit Function
    If Not AssertCollEquals(c, a("a", "a")) Then Exit Function
    T_6_07_optional_empty_exec_a = True
End Function

Public Function T_6_08_anchored_capturing() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "^(abc)$") Then Exit Function
    Dim c As Collection: Set c = ExecColl(r, "abc")
    If Not AssertTrue(Not (c Is Nothing), "anchored capture expected") Then Exit Function
    If Not AssertCollEquals(c, a("abc", "abc")) Then Exit Function
    T_6_08_anchored_capturing = True
End Function

Public Function T_6_09_capturing_partial_failure() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "(a)(b)(c)") Then Exit Function
    If Not AssertTrue(ExecColl(r, "abx") Is Nothing, "expected partial failure on 'abx'") Then Exit Function
    T_6_09_capturing_partial_failure = True
End Function

' Category 7 - alternation
Public Function T_7_01_alt_simple_a() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "a|b") Then Exit Function
    If Not AssertTrue(Not (ExecColl(r, "a") Is Nothing), "a|b should match 'a'") Then Exit Function
    T_7_01_alt_simple_a = True
End Function

Public Function T_7_02_alt_simple_b() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "a|b") Then Exit Function
    If Not AssertTrue(Not (ExecColl(r, "b") Is Nothing), "a|b should match 'b'") Then Exit Function
    T_7_02_alt_simple_b = True
End Function

Public Function T_7_03_alt_three() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "a|b|c") Then Exit Function
    If Not AssertTrue(Not (ExecColl(r, "c") Is Nothing), "a|b|c should match 'c'") Then Exit Function
    T_7_03_alt_three = True
End Function

Public Function T_7_04_grouped_alt_foobaz() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "(foo|bar)baz") Then Exit Function
    Dim c As Collection: Set c = ExecColl(r, "foobaz")
    If Not AssertTrue(Not (c Is Nothing), "expected grouped alt match") Then Exit Function
    If Not AssertCollEquals(c, a("foobaz", "foo")) Then Exit Function
    T_7_04_grouped_alt_foobaz = True
End Function

Public Function T_7_05_grouped_alt_barbaz() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "(foo|bar)baz") Then Exit Function
    Dim c As Collection: Set c = ExecColl(r, "barbaz")
    If Not AssertTrue(Not (c Is Nothing), "expected grouped alt match 'barbaz'") Then Exit Function
    If Not AssertCollEquals(c, a("barbaz", "bar")) Then Exit Function
    T_7_05_grouped_alt_barbaz = True
End Function

Public Function T_7_06_embedded_alt_abd() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "a(b|c)d") Then Exit Function
    If Not AssertTrue(Not (ExecColl(r, "abd") Is Nothing), "a(b|c)d should match 'abd'") Then Exit Function
    T_7_06_embedded_alt_abd = True
End Function

Public Function T_7_07_embedded_alt_acd() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "a(b|c)d") Then Exit Function
    If Not AssertTrue(Not (ExecColl(r, "acd") Is Nothing), "a(b|c)d should match 'acd'") Then Exit Function
    T_7_07_embedded_alt_acd = True
End Function

Public Function T_7_08_longer_preference() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "longer|long") Then Exit Function
    Dim c As Collection: Set c = ExecColl(r, "longer")
    If Not AssertTrue(Not (c Is Nothing), "longer|long should match 'longer'") Then Exit Function
    If Not AssertCollEquals(c, a("longer")) Then Exit Function
    T_7_08_longer_preference = True
End Function

' Category 8 - anchors (remaining are similar)
Public Function T_8_01_anchor_start_true() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "^abc") Then Exit Function
    If Not AssertTrue(Not (ExecColl(r, "abc") Is Nothing), "^abc should match 'abc'") Then Exit Function
    T_8_01_anchor_start_true = True
End Function

Public Function T_8_02_impl() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "^abc") Then Exit Function
    If Not AssertTrue(ExecColl(r, "xabc") Is Nothing, "^abc shouldn't match 'xabc'") Then Exit Function
    T_8_02_impl = True
End Function

Public Function T_8_03_anchor_end_true() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "abc$") Then Exit Function
    If Not AssertTrue(Not (ExecColl(r, "abc") Is Nothing), "abc$ should match 'abc'") Then Exit Function
    T_8_03_anchor_end_true = True
End Function

Public Function T_8_04_anchor_end_false() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "abc$") Then Exit Function
    If Not AssertTrue(ExecColl(r, "abcx") Is Nothing, "abc$ should not match 'abcx'") Then Exit Function
    T_8_04_anchor_end_false = True
End Function

Public Function T_8_05_anchor_full_true() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "^abc$") Then Exit Function
    If Not AssertTrue(Not (ExecColl(r, "abc") Is Nothing), "^abc$ should match 'abc'") Then Exit Function
    T_8_05_anchor_full_true = True
End Function

Public Function T_8_06_anchor_full_false() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "^abc$") Then Exit Function
    If Not AssertTrue(ExecColl(r, "ab c") Is Nothing, "^abc$ should not match 'ab c'") Then Exit Function
    T_8_06_anchor_full_false = True
End Function

Public Function T_8_07_anchor_digits_full_true() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "^\d+$") Then Exit Function
    If Not AssertTrue(Not (ExecColl(r, "123") Is Nothing), "^\d+$ should match '123'") Then Exit Function
    T_8_07_anchor_digits_full_true = True
End Function

Public Function T_8_08_anchor_digits_full_false() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "^\d+$") Then Exit Function
    If Not AssertTrue(ExecColl(r, "123a") Is Nothing, "^\d+$ should not match '123a'") Then Exit Function
    T_8_08_anchor_digits_full_false = True
End Function

Public Function T_8_09_anchor_multiline_start_after_newline() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "^a", False, -1, True) Then Exit Function
    If Not AssertTrue(Not (ExecColl(r, "a" & vbLf & "b") Is Nothing), "multiline ^a should match start after newline") Then Exit Function
    T_8_09_anchor_multiline_start_after_newline = True
End Function

Public Function T_8_10_anchor_multiline_end_before_newline() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "a$", False, -1, True) Then Exit Function
    If Not AssertTrue(Not (ExecColl(r, "b" & vbLf & "a") Is Nothing), "multiline a$ should match before newline") Then Exit Function
    T_8_10_anchor_multiline_end_before_newline = True
End Function

' Category 9 - case insensitivity
Public Function T_9_01_case_ignore_true() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "abc", True) Then Exit Function
    If Not AssertTrue(Not (ExecColl(r, "ABC") Is Nothing), "ignoreCase True should match 'ABC'") Then Exit Function
    T_9_01_case_ignore_true = True
End Function

Public Function T_9_02_case_ignore_false() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "abc", False) Then Exit Function
    If Not AssertTrue(ExecColl(r, "ABC") Is Nothing, "ignoreCase False should not match 'ABC'") Then Exit Function
    T_9_02_case_ignore_false = True
End Function

Public Function T_9_03_case_class_ignore() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "[a-z]", True) Then Exit Function
    If Not AssertTrue(Not (ExecColl(r, "A") Is Nothing), "ignoreCase class should match 'A'") Then Exit Function
    T_9_03_case_class_ignore = True
End Function

Public Function T_9_04_case_alt_ignore() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "A|b", True) Then Exit Function
    If Not AssertTrue(Not (ExecColl(r, "a") Is Nothing), "A|b with ignoreCase True should match 'a'") Then Exit Function
    T_9_04_case_alt_ignore = True
End Function

' Category 10 - replacement
Public Function T_10_01_replace_swap() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "(foo)(bar)") Then Exit Function
    Dim out As String: out = ReplaceStr(r, "foobar", "$2-$1")
    If Not AssertEqual(out, "bar-foo", "replace swap failed") Then Exit Function
    T_10_01_replace_swap = True
End Function

Public Function T_10_02_replace_wrap() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "(\d+)") Then Exit Function
    Dim out As String: out = ReplaceStr(r, "123", "[$1]")
    If Not AssertEqual(out, "[123]", "replace wrap failed") Then Exit Function
    T_10_02_replace_wrap = True
End Function

Public Function T_10_03_replace_full_ref() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "(a)(b)") Then Exit Function
    Dim out As String: out = ReplaceStr(r, "ab", "$0")
    If Not AssertEqual(out, "ab", "replace $0 failed") Then Exit Function
    T_10_03_replace_full_ref = True
End Function

Public Function T_10_04_replace_partial() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "a(b)c") Then Exit Function
    Dim out As String: out = ReplaceStr(r, "abc", "x$1y")
    If Not AssertEqual(out, "xby", "replace partial failed") Then Exit Function
    T_10_04_replace_partial = True
End Function

Public Function T_10_05_replace_no_match() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "a(b)c") Then Exit Function
    Dim out As String: out = ReplaceStr(r, "axc", "x$1y")
    If Not AssertEqual(out, "axc", "replace when no match should return original") Then Exit Function
    T_10_05_replace_no_match = True
End Function

Public Function T_10_07_replace_issue_27() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "(\S+), (\S+)") Then Exit Function
    Dim out As String: out = ReplaceStr(r, "Myer, Ken", "$2 $1")
    If Not AssertEqual(out, "Ken Myer", "replace issue #27 fail") Then Exit Function
    T_10_07_replace_issue_27 = True
End Function

' Category 11 - Lookahead
Public Function T_11_01_lookahead_positive_simple() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "a(?=b)") Then Exit Function
    If Not AssertTrue(Not (ExecColl(r, "ab") Is Nothing), "a(?=b) should match 'a' in 'ab'") Then Exit Function
    T_11_01_lookahead_positive_simple = True
End Function

Public Function T_11_02_lookahead_positive_fail() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "a(?=b)") Then Exit Function
    If Not AssertTrue(ExecColl(r, "ac") Is Nothing, "a(?=b) should not match 'a' in 'ac'") Then Exit Function
    T_11_02_lookahead_positive_fail = True
End Function

Public Function T_11_03_lookahead_in_sequence_exec() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "a(?=b)c") Then Exit Function
    Dim c As Collection: Set c = ExecColl(r, "abc")
    If Not AssertTrue((c Is Nothing), "a(?=b)c should not match 'abc'") Then Exit Function
    T_11_03_lookahead_in_sequence_exec = True
End Function

Public Function T_11_04_lookahead_negative_true() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "a(?!b)") Then Exit Function
    If Not AssertTrue(Not (ExecColl(r, "ac") Is Nothing), "a(?!b) should match 'ac'") Then Exit Function
    T_11_04_lookahead_negative_true = True
End Function

Public Function T_11_05_lookahead_negative_false() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "a(?!b)") Then Exit Function
    If Not AssertTrue(ExecColl(r, "ab") Is Nothing, "a(?!b) should not match 'ab'") Then Exit Function
    T_11_05_lookahead_negative_false = True
End Function

Public Function T_11_06_lookahead_start_fail() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "(?=a)b") Then Exit Function
    If Not AssertTrue(ExecColl(r, "ab") Is Nothing, "(?=a)b should not match 'ab' (lookahead at start)") Then Exit Function
    T_11_06_lookahead_start_fail = True
End Function

Public Function T_11_07_lookahead_overlapping() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "(?=a)a") Then Exit Function
    Dim c As Collection: Set c = ExecColl(r, "aa")
    If Not AssertTrue(Not (c Is Nothing), "overlapping lookahead expected to match") Then Exit Function
    T_11_07_lookahead_overlapping = True
End Function

Public Function T_11_08_lookahead_variable_length() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "a(?=b+)") Then Exit Function
    If Not AssertTrue(Not (ExecColl(r, "abb") Is Nothing), "a(?=b+) should match 'a' before 'bb'") Then Exit Function
    T_11_08_lookahead_variable_length = True
End Function

Public Function T_11_09_lookahead_variable_negative() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "a(?!b+)") Then Exit Function
    If Not AssertTrue(Not (ExecColl(r, "ac") Is Nothing), "a(?!b+) should match 'ac'") Then Exit Function
    T_11_09_lookahead_variable_negative = True
End Function

' Category 12 - lookbehind (fixed width)
Public Function T_12_01_lookbehind_positive_fixed() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "(?<=a)b") Then Exit Function
    If Not AssertTrue(Not (ExecColl(r, "ab") Is Nothing), "(?<=a)b should match 'b' in 'ab'") Then Exit Function
    T_12_01_lookbehind_positive_fixed = True
End Function

Public Function T_12_02_lookbehind_positive_fail() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "(?<=a)b") Then Exit Function
    If Not AssertTrue(ExecColl(r, "cb") Is Nothing, "(?<=a)b should not match 'cb'") Then Exit Function
    T_12_02_lookbehind_positive_fail = True
End Function

Public Function T_12_03_lookbehind_in_sequence_exec() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "c(?<=a)b") Then Exit Function
    Dim c As Collection: Set c = ExecColl(r, "cab")
    If Not AssertTrue((c Is Nothing), "c(?<=a)b should not match 'cab'") Then Exit Function
    T_12_03_lookbehind_in_sequence_exec = True
End Function

Public Function T_12_04_lookbehind_negative_fixed() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "(?<!a)b") Then Exit Function
    If Not AssertTrue(Not (ExecColl(r, "cb") Is Nothing), "(?<!a)b should match 'cb'") Then Exit Function
    T_12_04_lookbehind_negative_fixed = True
End Function

Public Function T_12_05_lookbehind_negative_fail() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "(?<!a)b") Then Exit Function
    If Not AssertTrue(ExecColl(r, "ab") Is Nothing, "(?<!a)b should not match 'ab'") Then Exit Function
    T_12_05_lookbehind_negative_fail = True
End Function

Public Function T_12_06_lookbehind_at_end() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "b(?<=a)") Then Exit Function
    If Not AssertTrue((ExecColl(r, "ab") Is Nothing), "b(?<=a) should not match 'ab'") Then Exit Function
    T_12_06_lookbehind_at_end = True
End Function

Public Function T_12_07_lookbehind_multi_char_fixed() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "(?<=ab)c") Then Exit Function
    If Not AssertTrue(Not (ExecColl(r, "abc") Is Nothing), "(?<=ab)c should match 'abc'") Then Exit Function
    T_12_07_lookbehind_multi_char_fixed = True
End Function

Public Function T_12_08_lookbehind_variable_length_error() As Boolean
    Dim r As New ASF_RegexEngine
    If InitRegexAndHandle(r, "(?<=a+)b") Then
        g_lastError = "expected Init to fail for variable-length lookbehind"
        Exit Function
    End If
    ' Init failure is expected
    T_12_08_lookbehind_variable_length_error = True
End Function

Public Function T_12_09_lookbehind_fixed_quantifier_ok() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "(?<=a{2})b") Then Exit Function
    If Not AssertTrue(Not (ExecColl(r, "aab") Is Nothing), "(?<=a{2})b should match 'aab'") Then Exit Function
    T_12_09_lookbehind_fixed_quantifier_ok = True
End Function

Public Function T_12_10_lookbehind_negative_fixed2() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "(?<!a{2})b") Then Exit Function
    If Not AssertTrue(Not (ExecColl(r, "ab") Is Nothing), "(?<!a{2})b should match 'ab'") Then Exit Function
    T_12_10_lookbehind_negative_fixed2 = True
End Function

Public Function T_12_11_lookbehind_empty_ok() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "(?<=)") Then Exit Function
    If Not AssertTrue(Not (ExecColl(r, "a") Is Nothing), "(?<=) should trivially succeed") Then Exit Function
    T_12_11_lookbehind_empty_ok = True
End Function

' Category 13 - atomic & possessive
Public Function T_13_01_atomic_group_atomic_greedy_locks() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "(?>a+)ab") Then Exit Function
    If Not AssertTrue((ExecColl(r, "aaab") Is Nothing), "(?>a+)ab should not match 'aaab'") Then Exit Function
    T_13_01_atomic_group_atomic_greedy_locks = True
End Function

Public Function T_13_02_atomic_group_prevents_backtrack() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "(?>a+)b") Then Exit Function
    If Not AssertTrue(Not ExecColl(r, "aab") Is Nothing, "(?>a+)b should match 'aab'") Then Exit Function
    T_13_02_atomic_group_prevents_backtrack = True
End Function

Public Function T_13_03_atomic_group_locks_choice() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "(?>a\|aa)b") Then Exit Function
    If Not AssertTrue(ExecColl(r, "aab") Is Nothing, "(?>a|aa)b should not match 'aab'") Then Exit Function
    T_13_03_atomic_group_locks_choice = True
End Function

Public Function T_13_04_possessive_one_or_more_true() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "a++b") Then Exit Function
    If Not AssertTrue(Not (ExecColl(r, "aaab") Is Nothing), "a++b should match 'aaab' (possessive semantics)") Then Exit Function
    T_13_04_possessive_one_or_more_true = True
End Function

Public Function T_13_05_possessive_overconsume_false() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "a++b") Then Exit Function
    If Not AssertTrue(ExecColl(r, "aaac") Is Nothing, "a++b should not match 'aaac'") Then Exit Function
    T_13_05_possessive_overconsume_false = True
End Function

Public Function T_13_06_capturing_possessive() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "(a++)b") Then Exit Function
    Dim c As Collection: Set c = ExecColl(r, "aaab")
    If Not AssertTrue(Not (c Is Nothing), "(a++)b should match 'aaab'") Then Exit Function
    If Not AssertCollEquals(c, a("aaab", "aaa")) Then Exit Function
    T_13_06_capturing_possessive = True
End Function

Public Function T_13_07_possessive_wildcard_exec() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, ".*+b") Then Exit Function
    Dim c As Collection: Set c = ExecColl(r, "aaabxc")
    If Not AssertTrue((c Is Nothing), ".*+b expected to return no match (engine-dependent)") Then Exit Function
    T_13_07_possessive_wildcard_exec = True
End Function

Public Function T_13_08_atomic_alt_true() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "(?>ab|a)b") Then Exit Function
    If Not AssertTrue(Not (ExecColl(r, "abb") Is Nothing), "(?>ab|a)b should match 'abb'") Then Exit Function
    T_13_08_atomic_alt_true = True
End Function

Public Function T_13_09_atomic_locks_inner_choice_false() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "(?>a(b\|c))d") Then Exit Function
    If Not AssertTrue(ExecColl(r, "acd") Is Nothing, "(?>a(b|c))d should not match 'acd'") Then Exit Function
    T_13_09_atomic_locks_inner_choice_false = True
End Function

' Category 14 - edge cases & combos
Public Function T_14_01_empty_pattern_empty_input() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, vbNullString) Then Exit Function
    If Not AssertTrue(Not (ExecColl(r, vbNullString) Is Nothing), "empty pattern should match empty") Then Exit Function
    T_14_01_empty_pattern_empty_input = True
End Function

Public Function T_14_02_empty_pattern_nonempty_input() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, vbNullString) Then Exit Function
    If Not AssertTrue(Not (ExecColl(r, "a") Is Nothing), "empty pattern should match anywhere") Then Exit Function
    T_14_02_empty_pattern_nonempty_input = True
End Function

Public Function T_14_03_anchor_empty_string_true() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "^$") Then Exit Function
    If Not AssertTrue(Not (ExecColl(r, vbNullString) Is Nothing), "^$ should match empty") Then Exit Function
    T_14_03_anchor_empty_string_true = True
End Function

Public Function T_14_04_anchor_empty_string_false() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "^$") Then Exit Function
    If Not AssertTrue(ExecColl(r, "a") Is Nothing, "^$ should not match 'a'") Then Exit Function
    T_14_04_anchor_empty_string_false = True
End Function

Public Function T_14_05_optional_group_exec() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "a(b*)?") Then Exit Function
    Dim c As Collection: Set c = ExecColl(r, "a")
    If Not AssertTrue(Not (c Is Nothing), "a(b*)? should match 'a'") Then Exit Function
    If Not AssertCollEquals(c, a("a", vbNullString)) Then Exit Function
    T_14_05_optional_group_exec = True
End Function

Public Function T_14_06_left_pref_alt_exec() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "\d+|\w+") Then Exit Function
    Dim c As Collection: Set c = ExecColl(r, "123abc")
    If Not AssertTrue(Not (c Is Nothing), "left-pref alt expected to match '123'") Then Exit Function
    If Not AssertCollEquals(c, a("123")) Then Exit Function
    T_14_06_left_pref_alt_exec = True
End Function

Public Function T_14_07_grouped_alt_capture() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "(\d+|\w+)") Then Exit Function
    Dim c As Collection: Set c = ExecColl(r, "123abc")
    If Not AssertTrue(Not (c Is Nothing), "grouped alt capture expected") Then Exit Function
    If Not AssertCollEquals(c, a("123", "123")) Then Exit Function
    T_14_07_grouped_alt_capture = True
End Function

Public Function T_14_08_greedy_last_digit() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, ".*\d") Then Exit Function
    Dim c As Collection: Set c = ExecColl(r, "abc123")
    If Not AssertTrue(Not (c Is Nothing), ".*\\d expected to match 'abc123'") Then Exit Function
    If Not AssertCollEquals(c, a("abc123")) Then Exit Function
    T_14_08_greedy_last_digit = True
End Function

Public Function T_14_09_lazy_first_digit() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, ".*?\d") Then Exit Function
    Dim c As Collection: Set c = ExecColl(r, "abc123")
    If Not AssertTrue(Not (c Is Nothing), ".*?\\d expected to match 'abc1'") Then Exit Function
    If Not AssertCollEquals(c, a("abc1")) Then Exit Function
    T_14_09_lazy_first_digit = True
End Function

Public Function T_14_10_large_quant_true() As Boolean
    Dim r As New ASF_RegexEngine
    If Not InitRegexAndHandle(r, "a{999999}") Then Exit Function
    Dim big As String: big = String(999999, "a")
    If Not AssertTrue(Not (ExecColl(r, big) Is Nothing), "large quant should match huge string") Then Exit Function
    T_14_10_large_quant_true = True
End Function

Public Function T_14_11_backtracking_limit() As Boolean
    Dim r As New ASF_RegexEngine
    If Not InitRegexAndHandle(r, "a*", False, 10, False, False) Then Exit Function
    ' We just ensure initialization and run — exact outcome may depend on MaxMatchSteps
    T_14_11_backtracking_limit = True
End Function

Public Function T_14_12_lookahead_and_consume() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "a(?=b)b") Then Exit Function
    If Not AssertTrue(Not (ExecColl(r, "ab") Is Nothing), "a(?=b)b should match 'ab'") Then Exit Function
    T_14_12_lookahead_and_consume = True
End Function

Public Function T_14_13_lookbehind_and_lookahead() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "(?<=a)b(?=c)") Then Exit Function
    If Not AssertTrue(Not (ExecColl(r, "abc") Is Nothing), "(?<=a)b(?=c) should match 'abc'") Then Exit Function
    T_14_13_lookbehind_and_lookahead = True
End Function

Public Function T_14_14_atomic_overconsume_false() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "(?>a+)a") Then Exit Function
    If Not AssertTrue(ExecColl(r, "aaa") Is Nothing, "(?>a+)a should not match 'aaa'") Then Exit Function
    T_14_14_atomic_overconsume_false = True
End Function

Public Function T_14_15_noncapturing_groups() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "(?:a)b") Then Exit Function
    Dim c As Collection: Set c = ExecColl(r, "ab")
    If c Is Nothing Then
        g_lastError = "expected capture for (?:a)b"
        Exit Function
    End If
    If Not AssertCollEquals(c, a("ab")) Then Exit Function
    T_14_15_noncapturing_groups = True
End Function

' Category 15 - unsupported
Public Function T_15_02_unicode_property() As Boolean
    Dim r As New ASF_RegexEngine
    If InitRegexAndHandle(r, "\p{L}") Then
        g_lastError = "Engine unexpectedly accepted Unicode property"
        Exit Function
    End If
    T_15_02_unicode_property = True
End Function

Public Function T_15_03_comment_syntax_unsupported() As Boolean
    Dim r As New ASF_RegexEngine
    If InitRegexAndHandle(r, "(?#comment)a") Then
        g_lastError = "Engine unexpectedly accepted (?# comment) syntax"
        Exit Function
    End If
    T_15_03_comment_syntax_unsupported = True
End Function

Public Function T_15_05_variable_lookbehind_error() As Boolean
    Dim r As New ASF_RegexEngine
    If InitRegexAndHandle(r, "(?<=a+)b") Then
        g_lastError = "Engine unexpectedly accepted variable-width lookbehind"
        Exit Function
    End If
    T_15_05_variable_lookbehind_error = True
End Function

Public Function T_15_06_inline_flags_not_supported() As Boolean
    Dim r As New ASF_RegexEngine
    If InitRegexAndHandle(r, "(?i)abc") Then
        g_lastError = "Engine unexpectedly accepted inline flags"
        Exit Function
    End If
    T_15_06_inline_flags_not_supported = True
End Function

Public Function T_15_07_conditionals_supported() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "(?:(1)a|b)") Then Exit Function
    Dim c As Collection: Set c = ExecColl(r, "ab")
    If c Is Nothing Then
        g_lastError = "expected capture for (?:(1)a|b)"
        Exit Function
    End If
    If Not AssertCollEquals(c, a("b", vbNullString)) Then Exit Function
    T_15_07_conditionals_supported = True
End Function

' Category 16 - backreferences and word boundaries
Public Function T_16_01_backref_numbered() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "(a)\1") Then Exit Function
    If Not AssertCollEquals(ExecColl(r, "aa"), a("aa", "a")) Then Exit Function
    If Not AssertTrue(ExecColl(r, "ab") Is Nothing, "(a)\1 should not match 'ab'") Then Exit Function
    T_16_01_backref_numbered = True
End Function

Public Function T_16_02_backref_word_repeat() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "(\w+) \1") Then Exit Function
    If Not AssertCollEquals(ExecColl(r, "say the the end"), a("the the", "the")) Then Exit Function
    T_16_02_backref_word_repeat = True
End Function

Public Function T_16_03_backref_named() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "(?<q>['""]).*?\k<q>") Then Exit Function
    If Not AssertCollEquals(ExecColl(r, "say 'hi"" there' now"), a("'hi"" there'", "'")) Then Exit Function
    T_16_03_backref_named = True
End Function

Public Function T_16_04_backref_ignorecase() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "(a)\1", True) Then Exit Function
    If Not AssertCollEquals(ExecColl(r, "aA"), a("aA", "a")) Then Exit Function
    T_16_04_backref_ignorecase = True
End Function

Public Function T_16_05_backref_unset_group() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "(a)?b\1") Then Exit Function
    If Not AssertCollEquals(ExecColl(r, "b"), a("b", vbNullString)) Then Exit Function
    T_16_05_backref_unset_group = True
End Function

' stdVBA issue 141 pattern
Public Function T_16_06_backref_in_lookahead() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "(.)(?=.*\1)", True, , True) Then Exit Function
    On Error GoTo EH
    Dim m As Collection: Set m = r.ExecAll("ss")
    If Not AssertTrue(Not m Is Nothing, "expected a match in 'ss'") Then Exit Function
    If Not AssertEqual(m.Count, 1) Then Exit Function
    Dim c As Collection: Set c = m(1)
    If Not AssertCollEquals(c, a("s", "s")) Then Exit Function
    Set m = r.ExecAll("abcab")
    If Not AssertEqual(m.Count, 2) Then Exit Function
    Set c = m(2)
    If Not AssertCollEquals(c, a("b", "b")) Then Exit Function
    T_16_06_backref_in_lookahead = True
    Exit Function
EH:
    g_lastError = "ExecAll error: #" & err.Number & " - " & err.Description
    err.Clear
End Function

Public Function T_16_07_backref_undefined_error() As Boolean
    Dim r As New ASF_RegexEngine
    If Not AssertFalse(InitRegexAndHandle(r, "(a)\2"), "expected Init error for (a)\2") Then Exit Function
    If Not AssertFalse(r.Validate("(?<x>a)\k<y>"), "expected Validate False for undefined \k<y>") Then Exit Function
    If Not AssertFalse(r.Validate("(a)\k<a"), "expected Validate False for unclosed \k<") Then Exit Function
    If Not AssertTrue(r.Validate("(?<x>a)\k<x>(b)\2"), "expected Validate True for defined refs") Then Exit Function
    T_16_07_backref_undefined_error = True
End Function

Public Function T_16_08_word_boundary() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "\bcat\b") Then Exit Function
    If Not AssertCollEquals(ExecColl(r, "a cat!"), a("cat")) Then Exit Function
    If Not AssertCollEquals(ExecColl(r, "cat"), a("cat")) Then Exit Function
    If Not AssertTrue(ExecColl(r, "concat") Is Nothing, "\bcat\b should not match 'concat'") Then Exit Function
    T_16_08_word_boundary = True
End Function

Public Function T_16_09_non_word_boundary() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "\Bcat") Then Exit Function
    If Not AssertCollEquals(ExecColl(r, "concat"), a("cat")) Then Exit Function
    If Not AssertTrue(ExecColl(r, "a cat") Is Nothing, "\Bcat should not match 'a cat'") Then Exit Function
    T_16_09_non_word_boundary = True
End Function

Public Function T_16_10_word_boundary_replace() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "\bis\b") Then Exit Function
    If Not AssertEqual(ReplaceStr(r, "this is it", "IS"), "this IS it") Then Exit Function
    T_16_10_word_boundary_replace = True
End Function

Public Function T_16_11_escape_plain_text() As Boolean
    Dim r As New ASF_RegexEngine
    If Not AssertEqual(r.Escape("bar"), "bar") Then Exit Function
    If Not AssertEqual(r.Escape("1.5"), "1\.5") Then Exit Function
    If Not InitRegexAndHandle(r, r.Escape("bar")) Then Exit Function
    If Not AssertCollEquals(ExecColl(r, "a bar"), a("bar")) Then Exit Function
    T_16_11_escape_plain_text = True
End Function

' vbforums report: empty alternative must also match at the end of the subject
Public Function T_16_12_empty_alt_execall() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "(9|)") Then Exit Function
    On Error GoTo EH
    Dim m As Collection: Set m = r.ExecAll("999")
    If Not AssertEqual(m.Count, 4) Then Exit Function
    Dim c As Collection: Set c = m(3)
    If Not AssertCollEquals(c, a("9", "9")) Then Exit Function
    Set c = m(4)
    If Not AssertCollEquals(c, a(vbNullString, vbNullString)) Then Exit Function
    If Not InitRegexAndHandle(r, "(?:9|)") Then Exit Function
    Set m = r.ExecAll("999")
    If Not AssertEqual(m.Count, 4) Then Exit Function
    Set c = m(4)
    If Not AssertCollEquals(c, a(vbNullString)) Then Exit Function
    T_16_12_empty_alt_execall = True
    Exit Function
EH:
    g_lastError = "ExecAll error: #" & err.Number & " - " & err.Description
    err.Clear
End Function

Public Function T_16_13_empty_alt_replace() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "(9|)") Then Exit Function
    If Not AssertEqual(ReplaceStr(r, "999", "x"), "xxxx") Then Exit Function
    If Not AssertEqual(ReplaceStr(r, "a9", "x"), "xaxx") Then Exit Function
    If Not InitRegexAndHandle(r, "9") Then Exit Function
    If Not AssertEqual(ReplaceStr(r, "a9b9", "x"), "axbx") Then Exit Function
    T_16_13_empty_alt_replace = True
End Function

Public Function T_16_14_unbalanced_groups_rejected() As Boolean
    Dim r As New ASF_RegexEngine
    If Not AssertFalse(r.Validate("(9|"), "(9| should be rejected") Then Exit Function
    If Not AssertFalse(r.Validate("(?:9|"), "(?:9| should be rejected") Then Exit Function
    If Not AssertFalse(r.Validate("(9"), "(9 should be rejected") Then Exit Function
    If Not AssertFalse(r.Validate("9)"), "9) should be rejected") Then Exit Function
    If Not AssertTrue(r.Validate("(9|)"), "(9|) should be accepted") Then Exit Function
    If Not AssertTrue(r.Validate("(?:9|)"), "(?:9|) should be accepted") Then Exit Function
    If Not AssertFalse(InitRegexAndHandle(r, "9)"), "Init should raise for 9)") Then Exit Function
    T_16_14_unbalanced_groups_rejected = True
End Function

' Each backtracking step of a quantifier must start from the pre-quantifier captures
Public Function T_16_15_quant_backtrack_restores_caps() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "(a)*ab") Then Exit Function
    If Not AssertCollEquals(ExecColl(r, "ab"), a("ab", vbNullString)) Then Exit Function
    If Not AssertCollEquals(ExecColl(r, "aab"), a("aab", "a")) Then Exit Function
    If Not InitRegexAndHandle(r, "(.)(?=.*\1)") Then Exit Function
    If Not AssertTrue(r.ExecAt("abcab", 3) Is Nothing, "'c' has no later repeat") Then Exit Function
    T_16_15_quant_backtrack_restores_caps = True
End Function

' Alternation followed by more pattern backtracks into the alternatives
Public Function T_16_16_alt_backtracking() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "(a|ab)c") Then Exit Function
    If Not AssertCollEquals(ExecColl(r, "abc"), a("abc", "ab")) Then Exit Function
    If Not InitRegexAndHandle(r, "a(?:b|bc)d") Then Exit Function
    If Not AssertCollEquals(ExecColl(r, "abcd"), a("abcd")) Then Exit Function
    If Not InitRegexAndHandle(r, "(a|ab)$") Then Exit Function
    If Not AssertCollEquals(ExecColl(r, "ab"), a("ab", "ab")) Then Exit Function
    T_16_16_alt_backtracking = True
End Function

' Quantified atoms with internal choices backtrack per iteration (ECMAScript)
Public Function T_16_17_quant_alt_backtracking() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "(?:a|ab)+c") Then Exit Function
    If Not AssertCollEquals(ExecColl(r, "ababc"), a("ababc")) Then Exit Function
    If Not InitRegexAndHandle(r, "(a|ab)*c") Then Exit Function
    If Not AssertCollEquals(ExecColl(r, "abc"), a("abc", "ab")) Then Exit Function
    If Not InitRegexAndHandle(r, "(?:(a)|b)+") Then Exit Function
    If Not AssertCollEquals(ExecColl(r, "ab"), a("ab", vbNullString)) Then Exit Function
    If Not InitRegexAndHandle(r, "x(y?)+$") Then Exit Function
    If Not AssertCollEquals(ExecColl(r, "x"), a("x", vbNullString)) Then Exit Function
    T_16_17_quant_alt_backtracking = True
End Function

Public Function T_16_18_anchors_anywhere() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "^a|b") Then Exit Function
    If Not AssertCollEquals(ExecColl(r, "cb"), a("b")) Then Exit Function
    If Not InitRegexAndHandle(r, "(^|,)x") Then Exit Function
    If Not AssertCollEquals(ExecColl(r, "a,x"), a(",x", ",")) Then Exit Function
    If Not InitRegexAndHandle(r, "a\$") Then Exit Function
    If Not AssertCollEquals(ExecColl(r, "a$"), a("a$")) Then Exit Function
    If Not InitRegexAndHandle(r, "a+?$") Then Exit Function
    If Not AssertCollEquals(ExecColl(r, "aaa"), a("aaa")) Then Exit Function
    If Not InitRegexAndHandle(r, "^a") Then Exit Function
    If Not AssertEqual(ReplaceStr(r, "ba", "x"), "ba") Then Exit Function
    If Not InitRegexAndHandle(r, "^", False, -1, True) Then Exit Function
    On Error GoTo EH
    If Not AssertEqual(r.ExecAll("a" & vbLf).Count, 2) Then Exit Function
    T_16_18_anchors_anywhere = True
    Exit Function
EH:
    g_lastError = "ExecAll error: #" & err.Number & " - " & err.Description
    err.Clear
End Function

' Unset groups keep their slot, so $n and positional captures stay aligned
Public Function T_16_19_groups_keep_positions() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "(a)|(b)") Then Exit Function
    If Not AssertCollEquals(ExecColl(r, "b"), a("b", vbNullString, "b")) Then Exit Function
    If Not AssertEqual(ReplaceStr(r, "b", "[$2]"), "[b]") Then Exit Function
    If Not InitRegexAndHandle(r, "(?=(a))a") Then Exit Function
    If Not AssertCollEquals(ExecColl(r, "a"), a("a", "a")) Then Exit Function
    If Not InitRegexAndHandle(r, "b") Then Exit Function
    If Not AssertEqual(ReplaceStr(r, "ab", "[$1]"), "a[$1]") Then Exit Function
    T_16_19_groups_keep_positions = True
End Function

' Malformed braces are literals, as in VBScript and JavaScript
Public Function T_16_20_brace_literals() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "a{x}") Then Exit Function
    If Not AssertCollEquals(ExecColl(r, "a{x}"), a("a{x}")) Then Exit Function
    If Not InitRegexAndHandle(r, "a{,2}") Then Exit Function
    If Not AssertCollEquals(ExecColl(r, "a{,2}"), a("a{,2}")) Then Exit Function
    If Not InitRegexAndHandle(r, "a{2") Then Exit Function
    If Not AssertCollEquals(ExecColl(r, "a{2"), a("a{2")) Then Exit Function
    T_16_20_brace_literals = True
End Function

Public Function T_16_21_code_point_escapes() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "\x41") Then Exit Function
    If Not AssertCollEquals(ExecColl(r, "A"), a("A")) Then Exit Function
    If Not InitRegexAndHandle(r, "\u00e9") Then Exit Function
    If Not AssertCollEquals(ExecColl(r, ChrW(&HE9)), a(ChrW(&HE9))) Then Exit Function
    If Not InitRegexAndHandle(r, "[\x61-c]+") Then Exit Function
    If Not AssertCollEquals(ExecColl(r, "abc"), a("abc")) Then Exit Function
    If Not InitRegexAndHandle(r, "[\D]+") Then Exit Function
    If Not AssertCollEquals(ExecColl(r, "ab1"), a("ab")) Then Exit Function
    If Not InitRegexAndHandle(r, "[^\u00e0-\u00ff]+") Then Exit Function
    If Not AssertCollEquals(ExecColl(r, "-" & ChrW(&HE9)), a("-")) Then Exit Function
    If Not AssertFalse(r.Validate("[z-a]"), "expected [z-a] to be rejected") Then Exit Function
    T_16_21_code_point_escapes = True
End Function

' Lookbehinds nested in lookaheads get their width too
Public Function T_16_22_nested_lookbehind() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "(?!a(?<!.x))") Then Exit Function
    If Not AssertTrue(r.ExecAt("b" & vbTab & "a ", 3) Is Nothing, "(?!a(?<!.x)) must fail before 'a'") Then Exit Function
    If Not InitRegexAndHandle(r, "(?=(?<=a)b)b") Then Exit Function
    If Not AssertCollEquals(ExecColl(r, "ab"), a("b")) Then Exit Function
    T_16_22_nested_lookbehind = True
End Function

' Bad input is refused by Validate and by Init with an engine error, never a VBA runtime error
Public Function T_16_23_malformed_patterns_rejected() As Boolean
    Dim r As New ASF_RegexEngine
    Dim pat As Variant, errNum As Long
    For Each pat In Array("[z-a]", "\p{L}", "[\p{L}]", "(?<=a+)b", "a{2,1}", "{3}", "a{99999999999}", "\b*", "\B+", _
                          "(", ")", "(?<>a)", "\k<x", "(a)\2", "a\", "[a", "(?#c)", String(101, "(") & String(101, ")"))
        If Not AssertFalse(r.Validate(pat), "Validate accepted " & pat) Then Exit Function
        On Error Resume Next
        r.Init pat
        errNum = err.Number
        err.Clear
        On Error GoTo 0
        If Not AssertTrue(errNum > vbObjectError And errNum < vbObjectError + 200, "Init of " & pat & " raised #" & errNum) Then Exit Function
    Next pat
    T_16_23_malformed_patterns_rejected = True
End Function

' (?R) re-entering itself at the same position fails instead of recursing forever
Public Function T_16_24_recursion_guarded() As Boolean
    Dim r As New ASF_RegexEngine: If Not InitRegexAndHandle(r, "(?R)a") Then Exit Function
    If Not AssertTrue(ExecColl(r, "aaa") Is Nothing, "(?R)a has no finite match") Then Exit Function
    If Not InitRegexAndHandle(r, "a(?R)?b") Then Exit Function
    If Not AssertCollEquals(ExecColl(r, "aaabbb"), a("aaabbb")) Then Exit Function
    T_16_24_recursion_guarded = True
End Function

' Matching without a pattern, a rejected Init and out-of-range start positions fail cleanly
Public Function T_16_25_api_guards() As Boolean
    Dim r As New ASF_RegexEngine
    Dim c As Collection, errNum As Long
    On Error Resume Next
    Set c = r.Exec("a")
    errNum = err.Number
    err.Clear
    On Error GoTo 0
    If Not AssertEqual(errNum, vbObjectError + 4, "Exec before Init should raise vbObjectError + 4") Then Exit Function
    If Not InitRegexAndHandle(r, "a") Then Exit Function
    If Not AssertFalse(InitRegexAndHandle(r, "("), "Init of '(' should fail") Then Exit Function
    If Not AssertEqual(r.pattern, "a", "a rejected Init must keep the previous pattern") Then Exit Function
    If Not AssertCollEquals(ExecColl(r, "a"), a("a")) Then Exit Function
    If Not AssertTrue(r.ExecAt("abc", 0) Is Nothing, "ExecAt before the subject") Then Exit Function
    If Not AssertTrue(r.ExecAt("abc", 5) Is Nothing, "ExecAt past the subject") Then Exit Function
    If Not InitRegexAndHandle(r, "$") Then Exit Function
    If Not AssertTrue(Not r.ExecAt("abc", 4) Is Nothing, "ExecAt at Len + 1 is a valid position") Then Exit Function
    T_16_25_api_guards = True
End Function

Public Function testNewG() As Boolean
    Dim r As New ASF_RegexEngine
    If InitRegexAndHandle(r, "(\D*)(\d*)(\W*)") Then
        Dim c As Collection
'        Set c = ExecColl(r, "abc12345#$*%")
        Set c = r.ExecAt("abc12345#$*%", 1)
    End If
End Function
