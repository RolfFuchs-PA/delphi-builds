"""
Converts PAApplications.bxp (SmartBear BuildStudio XML) to PowerShell.
"""
import xml.etree.ElementTree as ET
import re
import html
import sys

tree = ET.parse(r'C:\Work\BuildStudio\PAApplications.bxp')
root = tree.getroot()

out_lines = []
indent_level = 0
_skip_next_switch = False  # Set after project-selection radio group to skip the mapping switch

def emit(text=""):
    """Emit a line of PowerShell at the current indent level."""
    if text == "":
        out_lines.append("")
    else:
        out_lines.append("    " * indent_level + text)

def indent():
    global indent_level
    indent_level += 1

def dedent():
    global indent_level
    indent_level = max(0, indent_level - 1)

def get_prop(elem, prop_name, config_name="Release"):
    """Get a property value from an item's Properties/Configuration."""
    for cfg in elem.findall('./Properties/Configuration'):
        if cfg.get('Name', '') == config_name or config_name is None:
            for prop in cfg.findall('Property'):
                if prop.get('Name') == prop_name:
                    return prop.get('Value', '')
            # Also check nested properties
            for prop in cfg.findall('.//Property'):
                if prop.get('Name') == prop_name:
                    return prop.get('Value', '')
    return None

def get_nested_prop(elem, *path, config_name="Release"):
    """Get a nested property value following a path of property names."""
    for cfg in elem.findall('./Properties/Configuration'):
        if cfg.get('Name', '') == config_name or config_name is None:
            current = cfg
            for name in path:
                found = None
                for prop in current.findall('Property'):
                    if prop.get('Name') == name:
                        found = prop
                        break
                if found is None:
                    return None
                current = found
            return current.get('Value', '') if current is not None else None
    return None

def get_deep_prop(elem, prop_name, config_name="Release", parent_name=None):
    """Recursively search for a property value. If parent_name is given, only match under that parent."""
    for cfg in elem.findall('./Properties/Configuration'):
        if cfg.get('Name', '') == config_name or config_name is None:
            if parent_name:
                for prop in cfg.iter('Property'):
                    if prop.get('Name') == parent_name:
                        for child in prop.iter('Property'):
                            if child.get('Name') == prop_name:
                                return child.get('Value', '')
            else:
                for prop in cfg.iter('Property'):
                    if prop.get('Name') == prop_name:
                        return prop.get('Value', '')
    return None

def bxp_var_to_ps(text):
    """Convert %VAR% to $VAR in a string, using ${VAR} when followed by word chars."""
    if text is None:
        return ''
    # Unescape XML entities
    text = text.replace('&lt;', '<').replace('&gt;', '>').replace('&amp;', '&').replace('&quot;', '"')
    # Convert %VAR% to ${VAR} when followed by a word character, otherwise $VAR
    def _repl(m):
        var = m.group(1)
        after = text[m.end():m.end()+1] if m.end() < len(text) else ''
        if after and re.match(r'[A-Za-z0-9_]', after):
            return '${' + var + '}'
        return '$' + var
    text = re.sub(r'%([A-Za-z_][A-Za-z0-9_]*)%', _repl, text)
    return text

def ps_string(text):
    """Make a PowerShell double-quoted string, handling variable refs."""
    if text is None:
        return '""'
    t = bxp_var_to_ps(text)
    # Escape special PS chars except $ (for variable expansion)
    t = t.replace('`', '``').replace('"', '`"')
    # Escape colons after $VarName to avoid PS scope interpretation (e.g. $Revision: → $Revision`:)
    t = re.sub(r'(\$[A-Za-z_][A-Za-z0-9_]*):', r'\1`:', t)
    return f'"{t}"'

def ps_literal(text):
    """Make a PowerShell single-quoted string (no variable expansion)."""
    if text is None:
        return "''"
    t = text.replace('&lt;', '<').replace('&gt;', '>').replace('&amp;', '&').replace('&quot;', '"')
    t = t.replace("'", "''")
    return f"'{t}'"

def sanitize_func_name(name):
    """Create a valid PowerShell function name."""
    name = re.sub(r'[^A-Za-z0-9_-]', '', name.replace(' ', '-'))
    return name

def get_condition_ps(elem, config_name="Release"):
    """Convert a BuildStudio condition to PowerShell."""
    for cfg in elem.findall('./Properties/Configuration'):
        if cfg.get('Name', '') == config_name:
            cond_prop = None
            for prop in cfg.findall('Property'):
                if prop.get('Name') == 'Condition':
                    cond_prop = prop
                    break
            if cond_prop is None:
                return "$true"
            
            val1 = ''
            val2 = ''
            condition = ''
            operator = 'None'
            children_prop = None
            
            for sub in cond_prop.findall('Property'):
                n = sub.get('Name', '')
                v = sub.get('Value', '')
                if n == 'Value1': val1 = v
                elif n == 'Value2': val2 = v
                elif n == 'Condition': condition = v
                elif n == 'Operator': operator = v
                elif n == 'Children': children_prop = sub
            
            # Check for compound conditions (children items)
            child_items = []
            if children_prop is not None:
                child_items = children_prop.findall('Item')
            
            if child_items:
                # Compound condition: combine child conditions
                parts = []
                for child_item in child_items:
                    cv1 = ''
                    cv2 = ''
                    ccond = ''
                    cop = 'None'
                    for cp in child_item.findall('Property'):
                        cn = cp.get('Name', '')
                        cv = cp.get('Value', '')
                        if cn == 'Value1': cv1 = cv
                        elif cn == 'Value2': cv2 = cv
                        elif cn == 'Condition': ccond = cv
                        elif cn == 'Operator': cop = cv
                    
                    part = _format_single_condition(cv1, cv2, ccond)
                    if part:
                        parts.append((part, cop))
                
                if not parts:
                    return "$true"
                
                # Build combined expression
                result = parts[0][0]
                for i in range(1, len(parts)):
                    part_expr, prev_op = parts[i]
                    # The operator on the previous item determines how to join
                    join_op = parts[i-1][1]
                    if join_op == 'And':
                        result = f"({result}) -and ({part_expr})"
                    elif join_op == 'Or':
                        result = f"({result}) -or ({part_expr})"
                    else:
                        result = f"({result}) -and ({part_expr})"
                return result
            else:
                # Simple single condition
                result = _format_single_condition(val1, val2, condition)
                return result if result else "$true"
    return "$true"

def _format_single_condition(val1, val2, condition):
    """Format a single condition expression as PowerShell."""
    v1 = ps_string(val1)
    v2 = ps_string(val2)
    
    cond_map = {
        'oEQ': f'{v1} -ceq {v2}',
        'oEQIC': f'{v1} -eq {v2}',
        'oNOTEQ': f'{v1} -ne {v2}',
        'oNOTEQIC': f'{v1} -ne {v2}',
        'oNOTEQCS': f'{v1} -cne {v2}',
        'oGT': f'{v1} -gt {v2}',
        'oGTEQ': f'{v1} -ge {v2}',
        'oLT': f'{v1} -lt {v2}',
        'oLE': f'{v1} -le {v2}',
        'oLTEQ': f'{v1} -le {v2}',
        'oBI': f'{v1} -ne ""',
        'oBW': f'({v1}).StartsWith({v2})',
        'oBWIC': f'{v1} -like "{bxp_var_to_ps(val2)}*"',
        'oNOTBWIC': f'{v1} -notlike "{bxp_var_to_ps(val2)}*"',
        'oEW': f'({v1}).EndsWith({v2})',
        'oEWIC': f'{v1} -like "*{bxp_var_to_ps(val2)}"',
        'oNOTEWIC': f'{v1} -notlike "*{bxp_var_to_ps(val2)}"',
        'oCON': f'{v1} -clike "*{bxp_var_to_ps(val2)}*"',
        'oCONIC': f'{v1} -like "*{bxp_var_to_ps(val2)}*"',
        'oNOTCONIC': f'{v1} -notlike "*{bxp_var_to_ps(val2)}*"',
        'oContains': f'{v1} -like "*{bxp_var_to_ps(val2)}*"',
        'oContainsIC': f'{v1} -like "*{bxp_var_to_ps(val2)}*"',
        'oNotContains': f'{v1} -notlike "*{bxp_var_to_ps(val2)}*"',
        'oNotContainsIC': f'{v1} -notlike "*{bxp_var_to_ps(val2)}*"',
        'oStartsWith': f'({v1}).StartsWith({v2})',
        'oStartsWithIC': f'{v1} -like "{bxp_var_to_ps(val2)}*"',
        'oEndsWith': f'({v1}).EndsWith({v2})',
        'oEndsWithIC': f'{v1} -like "*{bxp_var_to_ps(val2)}"',
    }
    
    return cond_map.get(condition, f'# UNKNOWN CONDITION: {condition} - {v1} ?? {v2}')

def convert_delphiscript_to_ps(script_text, description=""):
    """Convert DelphiScript to PowerShell (best effort)."""
    if not script_text:
        return ["# Empty script"]
    
    lines = []
    lines.append(f"# Converted from DelphiScript: {description}")
    lines.append("# --- REVIEW THIS CONVERSION ---")
    
    # Clean up XML entities
    s = script_text.replace('&lt;', '<').replace('&gt;', '>').replace('&amp;', '&').replace('&quot;', '"')
    
    # Add each line as a comment with the original, then add a blank conversion
    for line in s.split('\n'):
        line = line.rstrip()
        if not line.strip():
            continue
        lines.append(f"# DELPHI: {line}")
    
    lines.append("# --- END DELPHISCRIPT (manual conversion required) ---")
    return lines

def process_item(item):
    """Process a single BuildStudio operation item and emit PowerShell."""
    name = item.get('Name', '')
    enabled = get_prop(item, 'Enabled')
    description = get_prop(item, 'Description') or ''
    
    if enabled == 'False':
        emit(f"# [DISABLED] {name}" + (f" - {description}" if description else ""))
        return
    
    op_name = re.sub(r'\s*\(.*\)$', '', name)  # Strip parenthesized description
    
    # Route to handler
    handler = HANDLERS.get(op_name, handle_unknown)
    handler(item, name, description)

def handle_comment(item, name, description):
    text = get_prop(item, 'Text') or description
    if text:
        for line in text.split('\n'):
            emit(f"# {line.rstrip()}")
    else:
        emit("# (comment)")

def handle_submacro_def(item, name, description):
    """Handle Submacro definition (function definition)."""
    emit("")
    emit(f"# === Submacro: {description} ===")
    
    # Get parameters
    params = []
    for cfg in item.findall('./Properties/Configuration'):
        for prop in cfg.findall('Property'):
            if prop.get('Name') == 'Parameters':
                for param_item in prop.findall('Item'):
                    pname = ''
                    ptype = 'In'
                    for pp in param_item.findall('Property'):
                        if pp.get('Name') == 'ParameterName':
                            pname = pp.get('Value', '')
                        if pp.get('Name') == 'ParameterType':
                            ptype = pp.get('Value', '')
                    if pname:
                        params.append((pname, ptype))
    
    func_name = sanitize_func_name(description or name)
    param_str = ', '.join([f'${p[0]}' for p in params])
    emit(f"function Invoke-{func_name} {{")
    if params:
        indent()
        emit("param(")
        indent()
        for i, (pname, ptype) in enumerate(params):
            comma = ',' if i < len(params) - 1 else ''
            emit(f"${pname}{comma}")
        dedent()
        emit(")")
        dedent()
    indent()
    emit("")
    
    # Process children
    children = item.find('Children')
    if children is not None:
        for child in children:
            process_item(child)
    
    dedent()
    emit("}")
    emit("")

def handle_initialization(item, name, description):
    emit("")
    emit("#" + "=" * 70)
    emit("# INITIALIZATION")
    emit("#" + "=" * 70)
    children = item.find('Children')
    if children is not None:
        for child in children:
            process_item(child)

def handle_group(item, name, description):
    emit("")
    emit(f"#region {description or name}")
    emit(f"Write-Log {ps_string(f'--- {description or name} ---')}")
    children = item.find('Children')
    if children is not None:
        for child in children:
            process_item(child)
    emit(f"#endregion {description or name}")
    emit("")

def handle_if_then(item, name, description):
    cond = get_condition_ps(item)
    desc_comment = f"  # {description}" if description else ""
    emit(f"if ({cond}) {{{desc_comment}")
    indent()
    children = item.find('Children')
    if children is not None:
        for child in children:
            process_item(child)
    dedent()
    emit("}")

def handle_else(item, name, description):
    desc_comment = f"  # {description}" if description else ""
    # Replace the last "}" with "} else {"
    if out_lines and out_lines[-1].rstrip() == "    " * (indent_level) + "}":
        out_lines[-1] = "    " * indent_level + f"}} else {{{desc_comment}"
    else:
        emit(f"else {{{desc_comment}")
    indent()
    children = item.find('Children')
    if children is not None:
        for child in children:
            process_item(child)
    dedent()
    emit("}")

def handle_set_variable(item, name, description):
    var_name = None
    var_value = None
    for cfg in item.findall('./Properties/Configuration'):
        for prop in cfg.findall('Property'):
            if prop.get('Name') == 'Variable':
                var_name = prop.get('Value', '')
            if prop.get('Name') == 'Value':
                var_value = prop.get('Value', '')
    if var_name and var_value is not None:
        desc_comment = f"  # {description}" if description else ""
        emit(f"${var_name} = {ps_string(var_value)}{desc_comment}")
    elif var_name:
        emit(f"${var_name} = ''  # {description}")
    else:
        emit(f"# TODO: Set variable - {description}")
    
    # Special: after DELPHI_VERSION is initialized to empty, inject detection logic
    if var_name == 'DELPHI_VERSION' and (var_value == '' or var_value is None):
        emit("# 1. Check INI for explicit DELPHI_VERSION setting")
        emit("$DELPHI_VERSION = Get-IniValue -Path \"${ABSOPENEDPROJECTDIR}PAApplications.ini\" -Section \"$INI_SECTION\" -Key \"DELPHI_VERSION\"")
        emit("# 2. Try to detect from source control path")
        emit("if (\"$DELPHI_VERSION\" -ceq \"\") {")
        indent()
        emit("if (\"$SOURCE_CONTROL_SOURCE_PATH\" -like \"*delphi 6*\") { $DELPHI_VERSION = \"6\" }")
        emit("if (\"$SOURCE_CONTROL_SOURCE_PATH\" -like \"*delphi 2007*\") { $DELPHI_VERSION = \"2007\" }")
        emit("if (\"$SOURCE_CONTROL_SOURCE_PATH\" -like \"*delphi xe2*\") { $DELPHI_VERSION = \"XE2\" }")
        emit("if (\"$SOURCE_CONTROL_SOURCE_PATH\" -like \"*delphi xe6*\") { $DELPHI_VERSION = \"XE6\" }")
        emit("if (\"$SOURCE_CONTROL_SOURCE_PATH\" -like \"*delphi 10.4*\" -or \"$SOURCE_CONTROL_SOURCE_PATH\" -like \"*sydney*\") { $DELPHI_VERSION = \"10.4\" }")
        dedent()
        emit("}")
        emit("# 3. Default to 10.4 Sydney if not detected")
        emit("if (\"$DELPHI_VERSION\" -ceq \"\") {")
        indent()
        emit("$DELPHI_VERSION = \"10.4\"")
        dedent()
        emit("}")
        emit("# 4. Always confirm with user -- pre-select the detected/default version")
        emit("$delphiOptions = @(\"Delphi 6\", \"Delphi 2007\", \"Delphi XE2\", \"Delphi XE6\", \"Delphi 10.4 Sydney\", \"Cancel build\")")
        emit("$delphiMap = @{ \"6\" = 0; \"2007\" = 1; \"XE2\" = 2; \"XE6\" = 3; \"10.4\" = 4 }")
        emit("$defaultIdx = $delphiMap[$DELPHI_VERSION] ?? 4")
        emit("Write-Log \"Delphi auto-detected as '$DELPHI_VERSION' for '$INI_SECTION' -- confirming with user...\"")
        emit("$delphiChoice = Show-RadioMenu -Title \"Confirm Delphi version for: $INI_SECTION\" -Options $delphiOptions -DefaultIndex $defaultIdx")
        emit("switch ($delphiChoice) {")
        indent()
        emit("0 { $DELPHI_VERSION = \"6\" }")
        emit("1 { $DELPHI_VERSION = \"2007\" }")
        emit("2 { $DELPHI_VERSION = \"XE2\" }")
        emit("3 { $DELPHI_VERSION = \"XE6\" }")
        emit("4 { $DELPHI_VERSION = \"10.4\" }")
        emit("default { throw \"Build cancelled - no Delphi version selected\" }")
        dedent()
        emit("}")
        emit("Write-Log \"Using Delphi version: $DELPHI_VERSION\"")

def handle_get_ini_value(item, name, description):
    ini_file = get_deep_prop(item, 'FileName')
    section = get_deep_prop(item, 'Section')
    key = get_deep_prop(item, 'ValueName')
    variable = get_deep_prop(item, 'ValueData')
    emit(f"${variable} = Get-IniValue -Path {ps_string(ini_file)} -Section {ps_string(section)} -Key {ps_string(key)}  # {description}")

def handle_set_ini_value(item, name, description):
    ini_file = get_deep_prop(item, 'FileName')
    section = get_deep_prop(item, 'Section')
    key = get_deep_prop(item, 'ValueName')
    value = get_deep_prop(item, 'ValueData')
    emit(f"Set-IniValue -Path {ps_string(ini_file)} -Section {ps_string(section)} -Key {ps_string(key)} -Value {ps_string(value)}  # {description}")

def handle_find_replace(item, name, description):
    filename = get_deep_prop(item, 'FileName')
    mode = get_deep_prop(item, 'Mode') or 'Find'
    variable = get_deep_prop(item, 'Variable')
    
    # Get find/replace items
    find_val = ''
    replace_val = ''
    for cfg in item.findall('./Properties/Configuration'):
        for prop in cfg.iter('Property'):
            if prop.get('Name') == 'FindReplace':
                for fr_item in prop.findall('Item'):
                    for fp in fr_item.iter('Property'):
                        if fp.get('Name') == 'Find':
                            for vp in fp.findall('Property'):
                                if vp.get('Name') == 'Value':
                                    find_val = vp.get('Value', '')
                        if fp.get('Name') == 'Replace':
                            for vp in fp.findall('Property'):
                                if vp.get('Name') == 'Value':
                                    replace_val = vp.get('Value', '')
    
    if mode == 'Find':
        emit(f"${variable} = Find-InFile -Path {ps_string(filename)} -Find {ps_string(find_val)}  # {description}")
    else:
        emit(f"Replace-InFile -Path {ps_string(filename)} -Find {ps_string(find_val)} -Replace {ps_string(replace_val)}  # {description}")

def handle_execute_dos(item, name, description):
    cmd = get_deep_prop(item, 'Command')
    working_dir = get_deep_prop(item, 'WorkingFolder')
    out_to_var = get_deep_prop(item, 'OutputResultsToVariable')
    variable = get_deep_prop(item, 'ResultsVariable')
    
    if out_to_var == 'True' and variable:
        if working_dir:
            emit(f"${variable} = Invoke-DosCommand -Command {ps_string(cmd)} -WorkingDirectory {ps_string(working_dir)}  # {description}")
        else:
            emit(f"${variable} = Invoke-DosCommand -Command {ps_string(cmd)}  # {description}")
    else:
        if working_dir:
            emit(f"Invoke-DosCommand -Command {ps_string(cmd)} -WorkingDirectory {ps_string(working_dir)}  # {description}")
        else:
            emit(f"Invoke-DosCommand -Command {ps_string(cmd)}  # {description}")

def handle_execute_program(item, name, description):
    program = get_deep_prop(item, 'FileName')
    params = get_deep_prop(item, 'Parameters')
    working_dir = get_deep_prop(item, 'WorkingFolder')
    exit_var = get_deep_prop(item, 'ExitCodeVariable')
    
    parts = [f"Invoke-Program -Path {ps_string(program)}"]
    if params:
        parts.append(f"-Arguments {ps_string(params)}")
    if working_dir:
        parts.append(f"-WorkingDirectory {ps_string(working_dir)}")
    if description:
        parts.append(f" # {description}")
    if exit_var:
        emit(f"${exit_var} = " + ' '.join(parts))
    else:
        emit(' '.join(parts))

def handle_copy_file(item, name, description):
    source = get_deep_prop(item, 'Source')
    dest = get_deep_prop(item, 'Destination')
    overwrite = get_deep_prop(item, 'OverwiteFiles')
    overwrite_ro = get_deep_prop(item, 'OverwiteReadOnly')
    include_sub = get_deep_prop(item, 'IncludeSubdirectories')
    dest_is_dir = get_deep_prop(item, 'DestinationIsDirectory')
    
    force = " -Force" if overwrite == 'True' or overwrite_ro == 'True' else ""
    recurse = " -Recurse" if include_sub == 'True' else ""
    
    emit(f"Copy-FileEx -Source {ps_string(source)} -Destination {ps_string(dest)}{force}{recurse}  # {description}")

def handle_move_file(item, name, description):
    source = get_deep_prop(item, 'Source')
    dest = get_deep_prop(item, 'Destination')
    emit(f"Move-Item -Path {ps_string(source)} -Destination {ps_string(dest)} -Force  # {description}")

def handle_delete_file(item, name, description):
    filename = get_deep_prop(item, 'FileName')
    include_sub = get_deep_prop(item, 'IncludeSubdirectories')
    recurse = " -Recurse" if include_sub == 'True' else ""
    emit(f"Remove-ItemSafe -Path {ps_string(filename)}{recurse}  # {description}")

def handle_write_to_file(item, name, description):
    filename = get_deep_prop(item, 'FileName')
    text = get_deep_prop(item, 'TextToWrite')
    write_type = get_deep_prop(item, 'WriteType')
    
    if write_type == 'Append':
        emit(f"Add-Content -Path {ps_string(filename)} -Value {ps_string(text)}  # {description}")
    else:
        emit(f"Set-Content -Path {ps_string(filename)} -Value {ps_string(text)}  # {description}")

def handle_read_from_file(item, name, description):
    filename = get_deep_prop(item, 'FileName')
    variable = get_deep_prop(item, 'OutputVariable')
    emit(f"${variable} = Get-Content -Path {ps_string(filename)} -Raw  # {description}")

def handle_if_file_exists(item, name, description):
    filename = get_deep_prop(item, 'FileName')
    variable = get_deep_prop(item, 'Variable')
    if variable:
        emit(f"${variable} = Test-Path {ps_string(filename)}  # {description}")
    else:
        emit(f"if (Test-Path {ps_string(filename)}) {{  # {description}")
        indent()
        children = item.find('Children')
        if children is not None:
            for child in children:
                process_item(child)
        dedent()
        emit("}")

def handle_if_dir_exists(item, name, description):
    dirname = get_deep_prop(item, 'DirectoryName')
    variable = get_deep_prop(item, 'Variable')
    if variable:
        emit(f"${variable} = Test-Path {ps_string(dirname)}  # {description}")
    else:
        emit(f"if (Test-Path {ps_string(dirname)}) {{  # {description}")
        indent()
        children = item.find('Children')
        if children is not None:
            for child in children:
                process_item(child)
        dedent()
        emit("}")

def handle_create_dir(item, name, description):
    dirname = get_deep_prop(item, 'Directory') or get_deep_prop(item, 'Path')
    emit(f"New-Item -ItemType Directory -Path {ps_string(dirname)} -Force | Out-Null  # {description}")

def handle_remove_dir(item, name, description):
    dirname = get_deep_prop(item, 'Directory') or get_deep_prop(item, 'Path')
    emit(f"Remove-ItemSafe -Path {ps_string(dirname)} -Recurse  # {description}")

def handle_rename_dir(item, name, description):
    source = get_deep_prop(item, 'SourceDirectory')
    dest = get_deep_prop(item, 'DestinationDirectory')
    emit(f"Rename-Item -Path {ps_string(source)} -NewName {ps_string(dest)} -Force  # {description}")

def handle_throw(item, name, description):
    msg = get_deep_prop(item, 'ExceptionMessage')
    if msg:
        emit(f"throw {ps_string(msg)}")
    else:
        emit(f"throw {ps_string(description or 'Build error')}")

def handle_log_message(item, name, description):
    text = get_deep_prop(item, 'Text') or get_deep_prop(item, 'Message')
    if text:
        emit(f"Write-Log {ps_string(text)}")
    elif description:
        emit(f"Write-Log {ps_string(description)}")

def handle_script(item, name, description):
    language = get_deep_prop(item, 'Language') or 'DelphiScript'
    script_text = get_deep_prop(item, 'ScriptText')
    
    emit(f"# Script block ({language}): {description}")
    lines = convert_delphiscript_to_ps(script_text, description)
    for line in lines:
        emit(line)

def handle_string_replace(item, name, description):
    # Properties: General.Input, General.Output, General.IsApplayToExistVal, Options.SearchStr, Options.ReplaceStr
    input_val = ''
    output_var = ''
    apply_to_existing = False
    search_str = ''
    replace_str = ''
    for cfg in item.findall('./Properties/Configuration'):
        for prop in cfg.findall('Property'):
            pn = prop.get('Name', '')
            if pn == 'General':
                for sp in prop.findall('Property'):
                    spn = sp.get('Name', '')
                    if spn == 'Input': input_val = sp.get('Value', '')
                    elif spn == 'Output': output_var = sp.get('Value', '')
                    elif spn == 'IsApplayToExistVal': apply_to_existing = sp.get('Value', '') == 'True'
            elif pn == 'Options':
                for sp in prop.findall('Property'):
                    spn = sp.get('Name', '')
                    if spn == 'SearchStr': search_str = sp.get('Value', '')
                    elif spn == 'ReplaceStr': replace_str = sp.get('Value', '')
    
    inp = ps_string(input_val)
    srch = bxp_var_to_ps(search_str) if search_str else ''
    repl = bxp_var_to_ps(replace_str) if replace_str else ''
    
    if apply_to_existing and input_val:
        # Modify the input variable in-place
        var_name = input_val.strip('%')
        emit(f"${var_name} = ${var_name}.Replace('{srch.replace(chr(39), chr(39)+chr(39))}', '{repl.replace(chr(39), chr(39)+chr(39))}')  # {description}")
    elif output_var:
        emit(f"${output_var} = ({inp}).Replace('{srch.replace(chr(39), chr(39)+chr(39))}', '{repl.replace(chr(39), chr(39)+chr(39))}')  # {description}")
    else:
        emit(f"# String Replace: {description} (search: {search_str} -> {replace_str})")

def handle_string_substring(item, name, description):
    input_val = ''
    output_var = ''
    start_str = ''
    end_str = ''
    apply_existing = False
    for cfg in item.findall('./Properties/Configuration'):
        for prop in cfg.findall('Property'):
            pn = prop.get('Name', '')
            if pn == 'General':
                for sp in prop.findall('Property'):
                    spn = sp.get('Name', '')
                    if spn == 'Input': input_val = sp.get('Value', '')
                    elif spn == 'Output': output_var = sp.get('Value', '')
                    elif spn == 'IsApplayToExistVal': apply_existing = sp.get('Value', '') == 'True'
            elif pn == 'Options':
                for sp in prop.findall('Property'):
                    spn = sp.get('Name', '')
                    if spn == 'StartStr': start_str = sp.get('Value', '')
                    elif spn == 'EndStr': end_str = sp.get('Value', '')
    
    target = input_val.strip('%') if apply_existing and input_val else output_var
    if target and start_str and end_str:
        emit(f"${target} = Get-SubstringBetween -Input {ps_string(input_val)} -Start {ps_string(start_str)} -End {ps_string(end_str)}  # {description}")
    elif target and start_str:
        emit(f"${target} = Get-SubstringAfter -Input {ps_string(input_val)} -Start {ps_string(start_str)}  # {description}")
    elif target:
        emit(f"# TODO: String Substring - check parameters for ${target}  # {description}")
    else:
        emit(f"# String Substring: {description}")

def handle_string_trimming(item, name, description):
    input_val = ''
    output_var = ''
    apply_existing = False
    for cfg in item.findall('./Properties/Configuration'):
        for prop in cfg.findall('Property'):
            pn = prop.get('Name', '')
            if pn == 'General':
                for sp in prop.findall('Property'):
                    spn = sp.get('Name', '')
                    if spn == 'Input': input_val = sp.get('Value', '')
                    elif spn == 'Output': output_var = sp.get('Value', '')
                    elif spn == 'IsApplayToExistVal': apply_existing = sp.get('Value', '') == 'True'
    
    target = input_val.strip('%') if apply_existing and input_val else output_var
    if target and input_val:
        emit(f"${target} = ({ps_string(input_val)}).Trim()  # {description}")
    else:
        emit(f"# String Trim: {description}")

def handle_string_concat(item, name, description):
    val1 = ''
    val2 = ''
    output_var = ''
    apply_existing = False
    for cfg in item.findall('./Properties/Configuration'):
        for prop in cfg.findall('Property'):
            pn = prop.get('Name', '')
            if pn == 'General':
                for sp in prop.findall('Property'):
                    spn = sp.get('Name', '')
                    if spn == 'Input': val1 = sp.get('Value', '')
                    elif spn == 'Output': output_var = sp.get('Value', '')
                    elif spn == 'IsApplayToExistVal': apply_existing = sp.get('Value', '') == 'True'
            elif pn == 'Options':
                for sp in prop.findall('Property'):
                    spn = sp.get('Name', '')
                    if spn == 'ConcatStr': val2 = sp.get('Value', '')
    
    target = val1.strip('%') if apply_existing and val1 else output_var
    if target:
        emit(f"${target} = {ps_string(val1)} + {ps_string(val2)}  # {description}")
    else:
        emit(f"# String Concatenation: {description}")

def handle_string_add_breaks(item, name, description):
    input_val = ''
    output_var = ''
    apply_existing = False
    for cfg in item.findall('./Properties/Configuration'):
        for prop in cfg.findall('Property'):
            pn = prop.get('Name', '')
            if pn == 'General':
                for sp in prop.findall('Property'):
                    spn = sp.get('Name', '')
                    if spn == 'Input': input_val = sp.get('Value', '')
                    elif spn == 'Output': output_var = sp.get('Value', '')
                    elif spn == 'IsApplayToExistVal': apply_existing = sp.get('Value', '') == 'True'
    
    target = input_val.strip('%') if apply_existing and input_val else output_var
    if target:
        emit(f"${target} = {ps_string(input_val)} + \"`r`n\"  # {description}")
    else:
        emit(f"# String Add Breaks: {description}")

def handle_path_manipulation(item, name, description):
    # Properties use nested General.Input, General.Output, Options.ExtractFileName, etc.
    input_val = ''
    output_var = ''
    apply_existing = False
    specify_output = False
    extract_filename = False
    extract_filepath = False
    change_ext = False
    remove_ext = False
    file_ext = ''
    
    for cfg in item.findall('./Properties/Configuration'):
        for prop in cfg.findall('Property'):
            pn = prop.get('Name', '')
            if pn == 'General':
                for sp in prop.findall('Property'):
                    spn = sp.get('Name', '')
                    if spn == 'Input': input_val = sp.get('Value', '')
                    elif spn == 'Output': output_var = sp.get('Value', '')
                    elif spn == 'IsApplayToExistVal': apply_existing = sp.get('Value', '') == 'True'
                    elif spn == 'IsSpecifyOutputVar': specify_output = sp.get('Value', '') == 'True'
            elif pn == 'Options':
                for sp in prop.findall('Property'):
                    spn = sp.get('Name', '')
                    if spn == 'ExtractFileName': extract_filename = sp.get('Value', '') == 'True'
                    elif spn == 'ExtractFilePath': extract_filepath = sp.get('Value', '') == 'True'
                    elif spn == 'ChangeFileExt': change_ext = sp.get('Value', '') == 'True'
                    elif spn == 'RemoveFileExt': remove_ext = sp.get('Value', '') == 'True'
                    elif spn == 'FileExt': file_ext = sp.get('Value', '')
    
    inp = ps_string(input_val)
    
    if apply_existing and input_val:
        target = input_val.strip('%')
    elif output_var:
        target = output_var
    else:
        target = None
    
    if target:
        if extract_filename:
            emit(f"${target} = Split-Path -Path {inp} -Leaf  # {description}")
        elif extract_filepath:
            emit(f"${target} = Split-Path -Path {inp} -Parent  # {description}")
        elif change_ext and file_ext:
            emit(f"${target} = [System.IO.Path]::ChangeExtension({inp}, '{bxp_var_to_ps(file_ext)}')  # {description}")
        elif remove_ext:
            emit(f"${target} = [System.IO.Path]::GetFileNameWithoutExtension({inp})  # {description}")
        else:
            emit(f"# Path Manipulation: {description}")
            emit(f"# ${target} = ... # TODO: verify path operation")
    else:
        emit(f"# Path Manipulation: {description}")

def handle_file_enumerator(item, name, description):
    file_mask = get_deep_prop(item, 'FileName')
    variable = get_deep_prop(item, 'FileNameVariable')
    count_var = get_deep_prop(item, 'CountVariable')
    include_sub = get_deep_prop(item, 'IncludeSubdirectories')
    
    # Split file mask into directory and filter
    recurse = " -Recurse" if include_sub == 'True' else ""
    
    vname = variable or 'item'
    if count_var:
        emit(f"${count_var} = 0")
    emit(f"foreach ($__file in (Get-ChildItem -Path {ps_string(file_mask)}{recurse} -ErrorAction SilentlyContinue)) {{  # {description}")
    indent()
    emit(f"${vname} = $__file.FullName")
    if count_var:
        emit(f"${count_var}++")
    children = item.find('Children')
    if children is not None:
        for child in children:
            process_item(child)
    dedent()
    emit("}")

def handle_file_content_enumerator(item, name, description):
    filename = get_deep_prop(item, 'FileName')
    variable = get_deep_prop(item, 'CurrentContElemVar') or 'CurrentLine'
    
    emit(f"foreach (${variable} in (Get-Content -Path {ps_string(filename)})) {{  # {description}")
    indent()
    children = item.find('Children')
    if children is not None:
        for child in children:
            process_item(child)
    dedent()
    emit("}")

def handle_try(item, name, description):
    emit(f"try {{  # {description}")
    indent()
    children = item.find('Children')
    if children is not None:
        for child in children:
            process_item(child)
    dedent()
    emit("}")

def handle_catch(item, name, description):
    if out_lines and out_lines[-1].rstrip().endswith("}"):
        out_lines[-1] = out_lines[-1].rstrip() + f" catch {{  # {description}"
    else:
        emit(f"catch {{  # {description}")
    indent()
    children = item.find('Children')
    if children is not None:
        for child in children:
            process_item(child)
    dedent()
    emit("}")

def handle_finally(item, name, description):
    if out_lines and out_lines[-1].rstrip().endswith("}"):
        out_lines[-1] = out_lines[-1].rstrip() + f" finally {{  # {description}"
    else:
        emit(f"finally {{  # {description}")
    indent()
    children = item.find('Children')
    if children is not None:
        for child in children:
            process_item(child)
    dedent()
    emit("}")

def handle_switch(item, name, description):
    global _skip_next_switch
    if _skip_next_switch:
        _skip_next_switch = False
        emit(f"# Switch: {description} (skipped — project selection handled by INI-based radio group above)")
        return
    emit(f"# Switch: {description}")
    # The switch variable is usually set before this
    switch_var = get_deep_prop(item, 'Variable') or get_deep_prop(item, 'Expression')
    emit(f"switch (${switch_var or 'VAR_RESULT'}) {{")
    indent()
    children = item.find('Children')
    if children is not None:
        for child in children:
            process_item(child)
    dedent()
    emit("}")

def handle_case(item, name, description):
    value = get_deep_prop(item, 'Value') or get_deep_prop(item, 'CaseValue')
    desc = description or value or ''
    emit(f"{ps_string(value)} {{  # {desc}")
    indent()
    children = item.find('Children')
    if children is not None:
        for child in children:
            process_item(child)
    dedent()
    emit("}")

def handle_break(item, name, description):
    emit("break")

def handle_stop_macro(item, name, description):
    emit(f"return  # Stop Macro Execution - {description}")

def handle_label(item, name, description):
    label = get_deep_prop(item, 'LabelName') or description
    emit(f"# LABEL: {label}  (GoTo not supported in PS - restructure logic)")

def handle_goto_label(item, name, description):
    label = get_deep_prop(item, 'LabelName') or description
    emit(f"# GOTO: {label}  (GoTo not supported in PS - restructure as loop)")

def handle_vault_checkout(item, name, description):
    repo = get_deep_prop(item, 'Repository')
    repo_folder = get_deep_prop(item, 'RepositoryFolder')
    filename = get_deep_prop(item, 'FileName')
    host = get_deep_prop(item, 'Host')
    username = get_deep_prop(item, 'Username')
    path = f"{repo_folder}/{filename}" if repo_folder and filename else repo_folder or ''
    emit(f"Invoke-VaultCheckOut -Repository {ps_string(repo)} -Path {ps_string(path)} -Host {ps_string(host)}  # {description}")

def handle_vault_get_latest(item, name, description):
    repo = get_deep_prop(item, 'Repository')
    repo_folder = get_deep_prop(item, 'RepositoryFolder')
    filename = get_deep_prop(item, 'FileName')
    local_folder = get_deep_prop(item, 'LocalFolder')
    use_default = get_deep_prop(item, 'UseDefaultWorkingFolder')
    path = f"{repo_folder}/{filename}" if repo_folder and filename else repo_folder or ''
    parts = [f"Invoke-VaultGetLatest -Repository {ps_string(repo)} -Path {ps_string(path)}"]
    if local_folder and use_default != 'True':
        parts.append(f"-LocalFolder {ps_string(local_folder)}")
    parts.append(f" # {description}")
    emit(' '.join(parts))

def handle_vault_checkin(item, name, description):
    repo = get_deep_prop(item, 'Repository')
    repo_folder = get_deep_prop(item, 'RepositoryFolder')
    filename = get_deep_prop(item, 'FileName')
    comment = get_deep_prop(item, 'Comment')
    path = f"{repo_folder}/{filename}" if repo_folder and filename else repo_folder or ''
    emit(f"Invoke-VaultCheckIn -Repository {ps_string(repo)} -Path {ps_string(path)} -Comment {ps_string(comment)}  # {description}")

def handle_vault_label(item, name, description):
    repo = get_deep_prop(item, 'Repository')
    repo_folder = get_deep_prop(item, 'RepositoryFolder')
    label = get_deep_prop(item, 'Label')
    path = repo_folder or ''
    emit(f"Invoke-VaultLabel -Repository {ps_string(repo)} -Path {ps_string(path)} -Label {ps_string(label)}  # {description}")

def handle_vault_undo_checkout(item, name, description):
    repo = get_deep_prop(item, 'Repository')
    repo_folder = get_deep_prop(item, 'RepositoryFolder')
    filename = get_deep_prop(item, 'FileName')
    path = f"{repo_folder}/{filename}" if repo_folder and filename else repo_folder or ''
    emit(f"Invoke-VaultUndoCheckOut -Repository {ps_string(repo)} -Path {ps_string(path)}  # {description}")

def handle_vault_custom(item, name, description):
    repo = get_deep_prop(item, 'Repository')
    cmd = get_deep_prop(item, 'Command')
    params = get_deep_prop(item, 'Parameters')
    emit(f"Invoke-VaultCommand -Repository {ps_string(repo)} -Command {ps_string(cmd)} -Parameters {ps_string(params)}  # {description}")

def handle_vault_file_enum(item, name, description):
    repo = get_deep_prop(item, 'Repository')
    repo_folder = get_deep_prop(item, 'RepositoryFolder')
    filename = get_deep_prop(item, 'FileName')
    variable = get_deep_prop(item, 'CurrentItemVariable') or 'vaultItem'
    path = repo_folder or ''
    
    emit(f"foreach (${variable} in (Get-VaultFiles -Repository {ps_string(repo)} -Path {ps_string(path)} -Filter {ps_string(filename)})) {{  # {description}")
    indent()
    children = item.find('Children')
    if children is not None:
        for child in children:
            process_item(child)
    dedent()
    emit("}")

def handle_vault_get_file(item, name, description):
    repo = get_deep_prop(item, 'Repository')
    repo_folder = get_deep_prop(item, 'RepositoryFolder')
    filename = get_deep_prop(item, 'FileName')
    label = get_deep_prop(item, 'Label')
    local_path = get_deep_prop(item, 'LocalFolder')
    path = f"{repo_folder}/{filename}" if repo_folder and filename else repo_folder or ''
    emit(f"Invoke-VaultGetByLabel -Repository {ps_string(repo)} -Path {ps_string(path)} -Label {ps_string(label)} -LocalPath {ps_string(local_path)}  # {description}")

def handle_get_xml_value(item, name, description):
    filename = get_deep_prop(item, 'FileName')
    xpath = get_deep_prop(item, 'XPath')
    variable = get_deep_prop(item, 'Variable') or get_deep_prop(item, 'OutputVariable')
    emit(f"${variable} = Get-XmlValue -Path {ps_string(filename)} -XPath {ps_string(xpath)}  # {description}")

def handle_set_xml_value(item, name, description):
    filename = get_deep_prop(item, 'FileName')
    xpath = get_deep_prop(item, 'XPath')
    value = get_deep_prop(item, 'Value')
    emit(f"Set-XmlValue -Path {ps_string(filename)} -XPath {ps_string(xpath)} -Value {ps_string(value)}  # {description}")

def handle_get_file_attrs(item, name, description):
    filename = get_deep_prop(item, 'FileName')
    modify_date_var = get_deep_prop(item, 'ModifyDateVar')
    creation_date_var = get_deep_prop(item, 'CreationDateVar')
    if modify_date_var:
        emit(f"${modify_date_var} = (Get-Item {ps_string(filename)}).LastWriteTime.ToString()  # {description}")
    if creation_date_var:
        emit(f"${creation_date_var} = (Get-Item {ps_string(filename)}).CreationTime.ToString()  # {description}")
    if not modify_date_var and not creation_date_var:
        emit(f"# Get File Attributes: {description}")

def handle_set_file_attrs(item, name, description):
    filename = get_deep_prop(item, 'FileName')
    emit(f"Set-ItemProperty -Path {ps_string(filename)} -Name IsReadOnly -Value $false  # {description}")

def handle_get_file_version(item, name, description):
    filename = get_deep_prop(item, 'FileName')
    
    # Extract keys from FixedVerInfoKeys Items
    key_vars = []
    for cfg in item.findall('./Properties/Configuration'):
        for prop in cfg.findall('Property'):
            if prop.get('Name') == 'FixedVerInfoKeys':
                for items_prop in prop.findall('Property'):
                    if items_prop.get('Name') == 'Items':
                        for key_item in items_prop.findall('Item'):
                            key_name = ''
                            var_name = ''
                            for kp in key_item.findall('Property'):
                                if kp.get('Name') == 'Key':
                                    key_name = kp.get('Value', '')
                                elif kp.get('Name') == 'Variable':
                                    var_name = kp.get('Value', '')
                            if key_name and var_name:
                                key_vars.append((key_name, var_name))
    
    if key_vars:
        emit(f"$__verInfo = (Get-Item {ps_string(filename)}).VersionInfo  # {description}")
        for key_name, var_name in key_vars:
            # Map BuildStudio key names to .NET VersionInfo properties
            key_map = {
                'FileVersion': 'FileVersion',
                'ProductVersion': 'ProductVersion',
                'FileDescription': 'FileDescription',
                'CompanyName': 'CompanyName',
                'ProductName': 'ProductName',
                'LegalCopyright': 'LegalCopyright',
            }
            prop_name = key_map.get(key_name, key_name)
            emit(f"${var_name} = $__verInfo.{prop_name}")
    else:
        emit(f"# Get File Version Info: {description} - no keys specified")

def handle_message_box(item, name, description):
    text = get_deep_prop(item, 'Text') or get_deep_prop(item, 'Message')
    emit(f"Write-Log \"[MESSAGE] {bxp_var_to_ps(text or description)}\"")

def handle_confirmation_box(item, name, description):
    question = get_deep_prop(item, 'Question')
    variable = get_deep_prop(item, 'QuesVariable') or 'VAR_RESULT'
    def_response = get_deep_prop(item, 'DefResponse')
    emit(f"${variable} = Confirm-Action -Message {ps_string(question or description)} -Default {ps_string(def_response or 'Yes')}  # {description}")

def handle_radio_group(item, name, description):
    variable = get_deep_prop(item, 'ResVar') or 'VAR_RESULT'
    title = get_deep_prop(item, 'QueryCaption') or description
    
    # Special case: project selection radio group — read from INI file dynamically
    if 'select project' in description.lower():
        emit(f"# Radio Group: {description} (dynamic from INI file)")
        emit(f"$iniSections = [System.Collections.Generic.List[string]]::new()")
        emit(f"foreach ($line in [System.IO.File]::ReadLines(\"${{ABSOPENEDPROJECTDIR}}PAApplications.ini\")) {{")
        indent()
        emit(f"if ($line -match '^\\[(.+)\\]$') {{ $iniSections.Add($Matches[1]) }}")
        dedent()
        emit(f"}}")
        emit(f"$projectOptions = @(\"None - cancel build\") + $iniSections.ToArray()")
        emit(f"${variable} = Show-RadioMenu -Title {ps_string(title)} -Options $projectOptions")
        emit(f"if (${variable} -eq 0) {{  # Cancel build")
        indent()
        emit(f"return")
        dedent()
        emit(f"}}")
        emit(f"$INI_SECTION = $projectOptions[${variable}]")
        global _skip_next_switch
        _skip_next_switch = True
        return
    
    # Extract options from List - items are directly under Property[@Name='List']/Item
    options = []
    for cfg in item.findall('./Properties/Configuration'):
        for prop in cfg.findall('Property'):
            if prop.get('Name') == 'List':
                for opt_item in prop.findall('Item'):
                    text = ''
                    index = ''
                    for op in opt_item.findall('Property'):
                        if op.get('Name') == 'ListItemText':
                            text = op.get('Value', '')
                        elif op.get('Name') == 'ListItemIndex':
                            index = op.get('Value', '')
                    if text:
                        options.append((text, index))
    
    emit(f"# Radio Group: {description}")
    emit(f"${variable} = Show-RadioMenu -Title {ps_string(title)} -Options @(")
    indent()
    for i, (text, idx) in enumerate(options):
        comma = ',' if i < len(options) - 1 else ''
        emit(f"{ps_string(text)}{comma}  # index={idx}")
    dedent()
    emit(")")

def handle_send_email(item, name, description):
    # Email sending removed — not needed in PowerShell version
    subject = get_deep_prop(item, 'Subject') or description
    emit(f"Write-Log \"[Email skipped] {bxp_var_to_ps(subject)}\"  # {description}")

def handle_export_log(item, name, description):
    filename = get_deep_prop(item, 'FileName')
    variable = get_deep_prop(item, 'Variable')
    export_mode = get_deep_prop(item, 'ExportMode')
    if variable:
        emit(f"${variable} = Get-BuildLog  # {description}")
    elif filename:
        emit(f"Export-BuildLog -Path {ps_string(filename)} -Mode {ps_string(export_mode or 'Text')}  # {description}")
    else:
        emit(f"Export-BuildLog  # {description}")

def handle_set_build_title(item, name, description):
    title = get_deep_prop(item, 'Title')
    emit(f"$script:BuildTitle = {ps_string(title)}  # {description}")
    emit(f"Write-Log \"Build Title: {bxp_var_to_ps(title or '')}\"")

def handle_create_guid(item, name, description):
    variable = get_deep_prop(item, 'OutputVar')
    guid_format = get_deep_prop(item, 'GUIDFormat') or 'B'
    emit(f"${variable} = [guid]::NewGuid().ToString('{guid_format}').ToUpper()  # {description}")

def handle_ini_file_enum(item, name, description):
    filename = get_deep_prop(item, 'FileName')
    section = get_deep_prop(item, 'Section')
    value_data_var = get_deep_prop(item, 'ValueDataVar') or 'CurrentValue'
    
    emit(f"foreach (${value_data_var} in (Get-IniSectionValues -Path {ps_string(filename)} -Section {ps_string(section)})) {{  # {description}")
    indent()
    children = item.find('Children')
    if children is not None:
        for child in children:
            process_item(child)
    dedent()
    emit("}")

def handle_compile_vs(item, name, description):
    solution = get_deep_prop(item, 'Solution')
    config = get_deep_prop(item, 'Configuration') or 'Release'
    compiler_ver = get_deep_prop(item, 'CompilerVersion')
    parts = [f"Invoke-MSBuild -SolutionFile {ps_string(solution)} -Configuration {ps_string(config)}"]
    if compiler_ver:
        parts.append(f"-CompilerVersion {ps_string(compiler_ver)}")
    parts.append(f" # {description}")
    emit(' '.join(parts))

def handle_installeware(item, name, description):
    project = get_deep_prop(item, 'WstrProjectFile')
    build_type = get_deep_prop(item, 'IBuildType')
    emit(f"Invoke-InstallAware -ProjectFile {ps_string(project)} -BuildType {ps_string(build_type or 'Release')}  # {description}")

def handle_if_prev_fails(item, name, description):
    emit(f"if (-not $?) {{  # {description}")
    indent()
    children = item.find('Children')
    if children is not None:
        for child in children:
            process_item(child)
    dedent()
    emit("}")

def handle_if_process_running(item, name, description):
    process_name = get_deep_prop(item, 'BstrProcessID')
    emit(f"if (Get-Process -Name {ps_string(process_name)} -ErrorAction SilentlyContinue) {{  # {description}")
    indent()
    children = item.find('Children')
    if children is not None:
        for child in children:
            process_item(child)
    dedent()
    emit("}")

def handle_terminate_process(item, name, description):
    process_name = get_deep_prop(item, 'BstrProcessID')
    emit(f"Stop-Process -Name {ps_string(process_name)} -Force -ErrorAction SilentlyContinue  # {description}")

def handle_run_submacro(item, name, description):
    submacro_name = get_deep_prop(item, 'Submacro')
    
    # Get parameter values
    params = []
    for cfg in item.findall('./Properties/Configuration'):
        for prop in cfg.findall('Property'):
            if prop.get('Name') == 'Parameters':
                for items_prop in prop.findall('Property'):
                    if items_prop.get('Name') == 'Items':
                        for param_item in items_prop.findall('Item'):
                            pname = ''
                            pval = ''
                            for pp in param_item.findall('Property'):
                                if pp.get('Name') == 'ParameterName':
                                    pname = pp.get('Value', '')
                                elif pp.get('Name') == 'ParameterValue':
                                    pval = pp.get('Value', '')
                            if pname:
                                params.append((pname, pval))
    
    func_name = sanitize_func_name(submacro_name or description or '')
    param_str = ' '.join([f"-{p[0]} {ps_string(p[1])}" for p in params])
    emit(f"Invoke-{func_name} {param_str}  # {description}")

def handle_copy_move_dir(item, name, description):
    source = get_deep_prop(item, 'Source')
    dest = get_deep_prop(item, 'Destination')
    emit(f"Copy-Item -Path {ps_string(source)} -Destination {ps_string(dest)} -Recurse -Force  # {description}")

def handle_if_xml_node_exists(item, name, description):
    xml_file = get_deep_prop(item, 'XMLFile')
    xpath = get_deep_prop(item, 'XPath')
    check_var = get_deep_prop(item, 'CheckResultVar')
    output_var = get_deep_prop(item, 'OutputVar')
    ret_type = get_deep_prop(item, 'RetValueType')
    fail_if_not = get_deep_prop(item, 'FailIfNodesNotExist')
    
    emit(f"# If XML Node/Attribute Exists: {description}")
    if output_var and ret_type == 'Get single node value':
        emit(f"${check_var or 'FIND_RESULT'} = 'False'")
        emit(f"${output_var} = ''")
        emit(f"try {{")
        indent()
        emit(f"[xml]$_xml = Get-Content -Path {ps_string(xml_file)} -Raw")
        emit(f"$_nsMgr = [System.Xml.XmlNamespaceManager]::new($_xml.NameTable)")
        emit(f"$_node = $_xml.SelectSingleNode({ps_string(xpath)}, $_nsMgr)")
        emit(f"if ($_node) {{")
        indent()
        emit(f"${check_var or 'FIND_RESULT'} = 'True'")
        emit(f"${output_var} = $_node.InnerText")
        dedent()
        emit(f"}}")
        dedent()
        emit(f"}} catch {{ Write-Log \"XML query failed: $_\" }}")
    else:
        emit(f"${check_var or 'FIND_RESULT'} = 'False'")
        emit(f"try {{")
        indent()
        emit(f"[xml]$_xml = Get-Content -Path {ps_string(xml_file)} -Raw")
        emit(f"$_nsMgr = [System.Xml.XmlNamespaceManager]::new($_xml.NameTable)")
        emit(f"$_node = $_xml.SelectSingleNode({ps_string(xpath)}, $_nsMgr)")
        emit(f"if ($_node) {{ ${check_var or 'FIND_RESULT'} = 'True' }}")
        dedent()
        emit(f"}} catch {{ Write-Log \"XML query failed: $_\" }}")
    
    if fail_if_not == 'True':
        emit(f"if (${check_var or 'FIND_RESULT'} -ne 'True') {{ throw \"XML node not found: {xpath}\" }}")
    
    children = item.find('Children')
    if children is not None:
        emit(f"if (${check_var or 'FIND_RESULT'} -ceq 'True') {{")
        indent()
        for child in children:
            process_item(child)
        dedent()
        emit("}")

def handle_string_reverse(item, name, description):
    input_val = get_deep_prop(item, 'Input', parent_name='General')
    apply_to_existing = get_deep_prop(item, 'IsApplayToExistVal', parent_name='General')
    output_var = get_deep_prop(item, 'OutputVar', parent_name='General')
    by_word = get_deep_prop(item, 'ByWord', parent_name='Options')
    
    if apply_to_existing == 'True':
        # Input is a variable reference like %VAR% — reverse in place
        var = bxp_var_to_ps(input_val) if input_val else '$TEMP_VAR_2'
        if by_word == 'True':
            emit(f"{var} = ({var} -split '\\s+')[-1..0] -join ' '  # String Reverse by word: {description}")
        else:
            emit(f"{var} = -join ({var}[-1..(-({var}).Length)])  # String Reverse: {description}")
    else:
        target = f"${output_var}" if output_var else '$TEMP_VAR_2'
        src = ps_string(input_val) if input_val else "''"
        if by_word == 'True':
            emit(f"{target} = ({src} -split '\\s+')[-1..0] -join ' '  # String Reverse by word: {description}")
        else:
            emit(f"{target} = -join ({src}[-1..(-({src}).Length)])  # String Reverse: {description}")

def handle_string_quoting(item, name, description):
    input_val = get_deep_prop(item, 'Input', parent_name='General')
    apply_to_existing = get_deep_prop(item, 'IsApplayToExistVal', parent_name='General')
    output_var = get_deep_prop(item, 'OutputVar', parent_name='General')
    add_double = get_deep_prop(item, 'AddDoubleQuotes', parent_name='Options')
    add_single = get_deep_prop(item, 'AddSingleQuotes', parent_name='Options')
    strip = get_deep_prop(item, 'StripQuotes', parent_name='Options')
    
    var = bxp_var_to_ps(input_val) if input_val else '$TEMP_VAR_2'
    target = var if apply_to_existing == 'True' else (f"${output_var}" if output_var else var)
    
    if strip == 'True':
        emit(f"{target} = {var}.Trim('\"', \"'\")  # String Quoting (strip): {description}")
    elif add_double == 'True':
        emit(f"{target} = '\"' + {var} + '\"'  # String Quoting (double): {description}")
    elif add_single == 'True':
        emit(f"{target} = \"'\" + {var} + \"'\"  # String Quoting (single): {description}")
    else:
        emit(f"# String Quoting: {description} (no-op)")

def handle_vault_cloak(item, name, description):
    repo = get_deep_prop(item, 'Repository')
    repo_folder = get_deep_prop(item, 'RepositoryFolder')
    emit(f"# Vault Cloak: {description}")
    emit(f"$auth = Get-VaultAuthArgs")
    emit(f"& $script:VaultExe cloak @auth -repository {ps_string(repo)} {ps_string(repo_folder)}")

def handle_vault_uncloak(item, name, description):
    repo = get_deep_prop(item, 'Repository')
    repo_folder = get_deep_prop(item, 'RepositoryFolder')
    emit(f"# Vault Uncloak: {description}")
    emit(f"$auth = Get-VaultAuthArgs")
    emit(f"& $script:VaultExe uncloak @auth -repository {ps_string(repo)} {ps_string(repo_folder)}")

def handle_edit_assembly_info(item, name, description):
    assemblies = get_deep_prop(item, 'Assemblies')
    emit(f"# Edit Assembly Info: {description}")
    emit(f"# Assembly file: {assemblies or 'unknown'}")
    emit(f"Write-Log \"TODO: Edit Assembly Info - {description} (currently not automated)\"")

def handle_click_window_button(item, name, description):
    emit(f"# Click Window Button: {description}")
    emit(f"# (Windows GUI automation - not applicable in PowerShell script)")
    emit(f"Write-Log \"[Skipped] Click Window Button: {description}\"")

def handle_wait_for_file(item, name, description):
    filename = get_deep_prop(item, 'FileName')
    wait_type = get_deep_prop(item, 'WaitForFileType')
    interval_str = get_deep_prop(item, 'Interval', parent_name='WaitOptions') or '60000'
    check_interval = get_deep_prop(item, 'CheckingInterval', parent_name='WaitOptions') or '2000'
    
    try:
        timeout_sec = int(interval_str) // 1000
    except (ValueError, TypeError):
        timeout_sec = 60
    try:
        poll_sec = int(check_interval) // 1000
    except (ValueError, TypeError):
        poll_sec = 2
    
    emit(f"# Wait for File: {description}")
    if wait_type == 'Exists':
        emit(f"$_waitEnd = (Get-Date).AddSeconds({timeout_sec})")
        emit(f"while (-not (Test-Path {ps_string(filename)}) -and (Get-Date) -lt $_waitEnd) {{")
        indent()
        emit(f"Start-Sleep -Seconds {poll_sec}")
        dedent()
        emit(f"}}")
    else:
        emit(f"Start-Sleep -Seconds {timeout_sec}  # Wait {timeout_sec}s")

def handle_unknown(item, name, description):
    emit(f"# TODO: [{name}] - {description}")
    children = item.find('Children')
    if children is not None and len(list(children)) > 0:
        indent()
        for child in children:
            process_item(child)
        dedent()

# Operation name -> handler mapping
HANDLERS = {
    'Comment': handle_comment,
    'Submacro': handle_submacro_def,
    'Initialization': handle_initialization,
    'Group': handle_group,
    'If ... Then': handle_if_then,
    'Else': handle_else,
    'Set/Reset Variable Value': handle_set_variable,
    'Get INI Value': handle_get_ini_value,
    'Set INI Value': handle_set_ini_value,
    'Find/Replace in File': handle_find_replace,
    'Execute DOS Command': handle_execute_dos,
    'Execute Program': handle_execute_program,
    'Copy File': handle_copy_file,
    'Copy File(s)': handle_copy_file,
    'Move File': handle_move_file,
    'Move File(s)': handle_move_file,
    'Delete File': handle_delete_file,
    'Delete File(s)': handle_delete_file,
    'Write to File': handle_write_to_file,
    'Read From File': handle_read_from_file,
    'If File Exists': handle_if_file_exists,
    'If Directory Exists': handle_if_dir_exists,
    'Create Directory': handle_create_dir,
    'Remove Directory': handle_remove_dir,
    'Rename Directory': handle_rename_dir,
    'Throw': handle_throw,
    'Log Message': handle_log_message,
    'Script': handle_script,
    'String Replace': handle_string_replace,
    'String Substring': handle_string_substring,
    'String Trimming': handle_string_trimming,
    'String Concatenation': handle_string_concat,
    'String Add Breaks': handle_string_add_breaks,
    'Path Manipulation': handle_path_manipulation,
    'File Enumerator': handle_file_enumerator,
    'File Content Enumerator': handle_file_content_enumerator,
    'Try': handle_try,
    'Catch': handle_catch,
    'Finally': handle_finally,
    'Switch': handle_switch,
    'Case': handle_case,
    'Break': handle_break,
    'Stop Macro Execution': handle_stop_macro,
    'Label': handle_label,
    'Go to Label': handle_goto_label,
    'Vault Check Out': handle_vault_checkout,
    'Vault Get Latest Version': handle_vault_get_latest,
    'Vault Check In': handle_vault_checkin,
    'Vault Label': handle_vault_label,
    'Vault Undo Check Out': handle_vault_undo_checkout,
    'Vault Custom Command': handle_vault_custom,
    'Vault File Enumerator': handle_vault_file_enum,
    'Vault Get File': handle_vault_get_file,
    'Vault Get File(s) By Label': handle_vault_get_file,
    'Get XML Value': handle_get_xml_value,
    'Set XML Value': handle_set_xml_value,
    'Get File Attributes': handle_get_file_attrs,
    'Set/Clear File Attributes': handle_set_file_attrs,
    'Get File Version Info': handle_get_file_version,
    'Message Box': handle_message_box,
    'Confirmation Box': handle_confirmation_box,
    'Radio Group': handle_radio_group,
    'Send E-mail': handle_send_email,
    'Export Log': handle_export_log,
    'Set Build Title': handle_set_build_title,
    'Create GUID': handle_create_guid,
    'INI File Enumerator': handle_ini_file_enum,
    'Compile Visual Studio Solution': handle_compile_vs,
    'InstallAware': handle_installeware,
    'If Previous Operation Fails': handle_if_prev_fails,
    'If Process is Running': handle_if_process_running,
    'Terminate Process': handle_terminate_process,
    'Run Submacro': handle_run_submacro,
    'Copy/Move Directory': handle_copy_move_dir,
    'If XML Node/Attribute Exists': handle_if_xml_node_exists,
    'String Reverse': handle_string_reverse,
    'String Quoting': handle_string_quoting,
    'Vault Cloak': handle_vault_cloak,
    'Vault Uncloak': handle_vault_uncloak,
    'Edit Assembly Info': handle_edit_assembly_info,
    'Click Window Button': handle_click_window_button,
    'Wait for File': handle_wait_for_file,
}

# ============================================================
# GENERATE THE POWERSHELL SCRIPT
# ============================================================

emit("#Requires -Version 7.0")
emit("<#")
emit(".SYNOPSIS")
emit("    PA Applications Build Script")
emit("    Converted from SmartBear BuildStudio (PAApplications.bxp)")
emit("")
emit(".DESCRIPTION")
emit("    Builds PA Applications Delphi projects, runs tests, builds setups,")
emit("    and manages Vault source control operations.")
emit("")
emit(".PARAMETER INI_SECTION")
emit("    Project name to build (e.g., 'Bank Reconciliation', 'JET').")
emit("    If not provided, interactive selection is shown.")
emit("")
emit(".PARAMETER NIGHTLY_BUILD")
emit("    Set to TRUE for nightly/unattended builds.")
emit("#>")
emit("[CmdletBinding()]")
emit("param(")
emit("    [string]$INI_SECTION = '',")
emit("    [string]$NIGHTLY_BUILD = 'FALSE'")
emit(")")
emit("")
emit("Set-StrictMode -Version Latest")
emit("$ErrorActionPreference = 'Stop'")
emit("")
emit("# BuildStudio built-in: directory containing this script (with trailing backslash)")
emit("$ABSOPENEDPROJECTDIR = $PSScriptRoot + [IO.Path]::DirectorySeparatorChar")
emit("")

# ============================================================
# CONSTANTS
# ============================================================
emit("#" + "=" * 70)
emit("# CONSTANTS")
emit("#" + "=" * 70)
for item in root.findall('.//Constants/Item'):
    desc = item.get('Description', '')
    for cfg in item.findall('Configuration'):
        for const in cfg.findall('Constant'):
            cname = const.get('Name', '')
            cval = const.get('Value', '') or const.get('DefaultValue', '')
            comment = f"  # {desc}" if desc else ""
            emit(f"$script:{cname} = '{cval}'{comment}")
emit("")

# ============================================================
# VARIABLES
# ============================================================
emit("#" + "=" * 70)
emit("# VARIABLES")
emit("#" + "=" * 70)
for item in root.findall('.//Variables/Item'):
    desc = item.get('Description', '')
    for cfg in item.findall('.//Variable'):
        vname = cfg.get('Name', '')
        vval = cfg.get('DefaultValue', '')
        comment = f"  # {desc}" if desc else ""
        if vname in ('INI_SECTION', 'NIGHTLY_BUILD'):
            continue  # These are params
        emit(f"${vname} = '{vval}'{comment}")
emit("")

# Vault credentials
emit("#" + "-" * 70)
emit("# Vault Source Control Credentials")
emit("#" + "-" * 70)
emit("$VAULT_USERNAME = $env:VAULT_USERNAME ?? 'autobuild'")
emit("$VAULT_PASSWORD = $env:VAULT_PASSWORD ?? 'autobuild'")
emit("")

# ============================================================
# HELPER FUNCTIONS
# ============================================================
emit("#" + "=" * 70)
emit("# HELPER FUNCTIONS")
emit("#" + "=" * 70)
emit("")

# Write-Log
emit("function Write-Log {")
emit("    param([string]$Message)")
emit("    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'")
emit("    Write-Host \"[$timestamp] $Message\"")
emit("    if ($script:LogFile) {")
emit("        \"[$timestamp] $Message\" >> $script:LogFile")
emit("    }")
emit("}")
emit("")

# Get-IniValue
emit("function Get-IniValue {")
emit("    param([string]$Path, [string]$Section, [string]$Key)")
emit("    if (-not (Test-Path $Path)) { return '' }")
emit("    $inSection = $false")
emit("    foreach ($line in [System.IO.File]::ReadLines($Path)) {")
emit("        if ($line -match '^\\[(.+)\\]$') {")
emit("            $inSection = ($Matches[1] -eq $Section)")
emit("        } elseif ($inSection -and $line -match \"^$([regex]::Escape($Key))\\s*=\\s*(.*)\") {")
emit("            $raw = $Matches[1].Trim()")
emit("            # Expand %VAR% references using current script variables")
emit("            $expanded = [regex]::Replace($raw, '%([A-Za-z_][A-Za-z0-9_]*)%', {")
emit("                param($m)")
emit("                $v = Get-Variable -Name $m.Groups[1].Value -ValueOnly -ErrorAction SilentlyContinue")
emit("                if ($null -ne $v) { $v } else { $m.Value }")
emit("            })")
emit("            return $expanded")
emit("        }")
emit("    }")
emit("    return ''")
emit("}")
emit("")

# Set-IniValue
emit("function Set-IniValue {")
emit("    param([string]$Path, [string]$Section, [string]$Key, [string]$Value)")
emit("    if (-not (Test-Path $Path)) {")
emit("        Set-Content -Path $Path -Value \"[$Section]`r`n$Key=$Value\"")
emit("        return")
emit("    }")
emit("    $lines = [System.IO.File]::ReadAllLines($Path)")
emit("    $result = [System.Collections.Generic.List[string]]::new($lines.Length + 2)")
emit("    $inSection = $false")
emit("    $keyFound = $false")
emit("    foreach ($line in $lines) {")
emit("        if ($line -match '^\\[(.+)\\]$') {")
emit("            if ($inSection -and -not $keyFound) {")
emit("                $result.Add(\"$Key=$Value\")")
emit("                $keyFound = $true")
emit("            }")
emit("            $inSection = ($Matches[1] -eq $Section)")
emit("        } elseif ($inSection -and $line -match \"^$([regex]::Escape($Key))\\s*=\") {")
emit("            $result.Add(\"$Key=$Value\")")
emit("            $keyFound = $true")
emit("            continue")
emit("        }")
emit("        $result.Add($line)")
emit("    }")
emit("    if (-not $keyFound) {")
emit("        if (-not $inSection) { $result.Add(\"[$Section]\") }")
emit("        $result.Add(\"$Key=$Value\")")
emit("    }")
emit("    [System.IO.File]::WriteAllLines($Path, $result)")
emit("}")
emit("")

# Get-IniSectionKeys
emit("function Get-IniSectionKeys {")
emit("    param([string]$Path, [string]$Section)")
emit("    $keys = [System.Collections.Generic.List[string]]::new()")
emit("    if (-not (Test-Path $Path)) { return $keys }")
emit("    $inSection = $false")
emit("    foreach ($line in [System.IO.File]::ReadLines($Path)) {")
emit("        if ($line -match '^\\[(.+)\\]$') {")
emit("            $inSection = ($Matches[1] -eq $Section)")
emit("        } elseif ($inSection -and $line -match '^(.+?)\\s*=') {")
emit("            $keys.Add($Matches[1])")
emit("        }")
emit("    }")
emit("    return $keys")
emit("}")
emit("")

# Get-IniSectionValues (for INI File Enumerator)
emit("function Get-IniSectionValues {")
emit("    param([string]$Path, [string]$Section)")
emit("    $values = [System.Collections.Generic.List[string]]::new()")
emit("    if (-not (Test-Path $Path)) { return $values }")
emit("    $inSection = $false")
emit("    foreach ($line in [System.IO.File]::ReadLines($Path)) {")
emit("        if ($line -match '^\\[(.+)\\]$') {")
emit("            $inSection = ($Matches[1] -eq $Section)")
emit("        } elseif ($inSection -and $line -match '^.+?\\s*=\\s*(.*)') {")
emit("            $values.Add($Matches[1].Trim())")
emit("        }")
emit("    }")
emit("    return $values")
emit("}")
emit("")

# Find-InFile
emit("function Find-InFile {")
emit("    param([string]$Path, [string]$Find)")
emit("    if (-not (Test-Path $Path)) { return '' }")
emit("    $content = Get-Content -Path $Path -Raw")
emit("    if ($content -match [regex]::Escape($Find)) { return $Find }")
emit("    return ''")
emit("}")
emit("")

# Replace-InFile
emit("function Replace-InFile {")
emit("    param([string]$Path, [string]$Find, [string]$Replace)")
emit("    if (-not (Test-Path $Path)) { return }")
emit("    $content = Get-Content -Path $Path -Raw")
emit("    $content = $content.Replace($Find, $Replace)")
emit("    Set-Content -Path $Path -Value $content -NoNewline")
emit("}")
emit("")

# Get-SubstringBetween
emit("function Get-SubstringBetween {")
emit("    param([string]$Input, [string]$Start, [string]$End)")
emit("    $startIdx = $Input.IndexOf($Start)")
emit("    if ($startIdx -lt 0) { return '' }")
emit("    $startIdx += $Start.Length")
emit("    $endIdx = $Input.IndexOf($End, $startIdx)")
emit("    if ($endIdx -lt 0) { return $Input.Substring($startIdx) }")
emit("    return $Input.Substring($startIdx, $endIdx - $startIdx)")
emit("}")
emit("")

# Get-SubstringAfter
emit("function Get-SubstringAfter {")
emit("    param([string]$Input, [string]$Start)")
emit("    $idx = $Input.IndexOf($Start)")
emit("    if ($idx -lt 0) { return '' }")
emit("    return $Input.Substring($idx + $Start.Length)")
emit("}")
emit("")

# Remove-ItemSafe
emit("function Remove-ItemSafe {")
emit("    param([string]$Path, [switch]$Recurse)")
emit("    if (Test-Path $Path) {")
emit("        if ($Recurse) { Remove-Item -Path $Path -Recurse -Force -ErrorAction SilentlyContinue }")
emit("        else { Remove-Item -Path $Path -Force -ErrorAction SilentlyContinue }")
emit("    }")
emit("}")
emit("")

# Copy-FileEx (handles wildcards)
emit("function Copy-FileEx {")
emit("    param([string]$Source, [string]$Destination, [switch]$Force, [switch]$Recurse)")
emit("    $srcDir = Split-Path -Path $Source -Parent")
emit("    $srcFilter = Split-Path -Path $Source -Leaf")
emit("    if (-not (Test-Path $srcDir)) {")
emit("        Write-Log \"[WARNING] Source directory not found: $srcDir\"")
emit("        return $false")
emit("    }")
emit("    if ($Recurse) {")
emit("        Copy-Item -Path $Source -Destination $Destination -Force:$Force -Recurse -ErrorAction Stop")
emit("    } else {")
emit("        Copy-Item -Path $Source -Destination $Destination -Force:$Force -ErrorAction Stop")
emit("    }")
emit("    return $true")
emit("}")
emit("")

# Invoke-DosCommand
emit("function Invoke-DosCommand {")
emit("    param([string]$Command, [string]$WorkingDirectory)")
emit("    Write-Log \"Executing: $Command\"")
emit("    $origDir = $PWD")
emit("    try {")
emit("        if ($WorkingDirectory -and (Test-Path $WorkingDirectory)) {")
emit("            Set-Location $WorkingDirectory")
emit("        }")
emit("        $output = cmd.exe /c $Command 2>&1")
emit("        $script:LastExitCode = $LASTEXITCODE")
emit("        return ($output -join \"`r`n\")")
emit("    } finally {")
emit("        Set-Location $origDir")
emit("    }")
emit("}")
emit("")

# Invoke-Program
emit("function Invoke-Program {")
emit("    param([string]$Path, [string]$Arguments, [string]$WorkingDirectory, [int]$TimeoutSeconds = 0)")
emit("    Write-Log \"Running: $Path $Arguments\"")
emit("    $psi = [System.Diagnostics.ProcessStartInfo]@{")
emit("        FileName               = $Path")
emit("        Arguments              = $Arguments ?? ''")
emit("        WorkingDirectory       = $WorkingDirectory ?? $PWD.Path")
emit("        UseShellExecute        = $false")
emit("        RedirectStandardOutput = $true")
emit("        RedirectStandardError  = $true")
emit("    }")
emit("    $process = [System.Diagnostics.Process]::Start($psi)")
emit("    $stdOut = $process.StandardOutput.ReadToEndAsync()")
emit("    $stdErr = $process.StandardError.ReadToEndAsync()")
emit("    $exited = if ($TimeoutSeconds -gt 0) {")
emit("        $process.WaitForExit($TimeoutSeconds * 1000)")
emit("    } else {")
emit("        $process.WaitForExit(); $true")
emit("    }")
emit("    if (-not $exited) {")
emit("        $process.Kill($true)")
emit("        throw \"Process timed out after ${TimeoutSeconds}s: $Path\"")
emit("    }")
emit("    [System.Threading.Tasks.Task]::WaitAll($stdOut, $stdErr)")
emit("    $script:LAST_STDOUT = $stdOut.Result")
emit("    $script:LAST_STDERR = $stdErr.Result")
emit("    return $process.ExitCode")
emit("}")
emit("")

# Vault helper functions
emit("#region Vault Source Control Helpers")
emit("")
emit("$script:VaultExe = 'C:\\Program Files (x86)\\SourceGear\\Vault Client\\vault.exe'")
emit("if (-not (Test-Path $script:VaultExe)) {")
emit("    $found = Get-Command vault.exe -ErrorAction SilentlyContinue")
emit("    $script:VaultExe = $found ? $found.Source : 'vault.exe'")
emit("}")
emit("")
emit("function Get-VaultAuthArgs {")
emit("    $authArgs = @('-host', $VAULT_SERVER_ADDRESS, '-user', $VAULT_USERNAME)")
emit("    if ($VAULT_PASSWORD) { $authArgs += @('-password', $VAULT_PASSWORD) }")
emit("    return $authArgs")
emit("}")
emit("")
emit("function Get-VaultWorkingFolder {")
emit("    param([string]$Repository, [string]$ReposPath)")
emit("    $auth = Get-VaultAuthArgs")
emit("    $output = (& $script:VaultExe listworkingfolders @auth -repository $Repository 2>$null) -join \"`n\"")
emit("    try {")
emit("        $xml = [xml]$output")
emit("        $wf = $xml.vault.listworkingfolders.workingfolder | Where-Object { $_.reposfolder -eq $ReposPath }")
emit("        return $wf ? $wf.localfolder : $null")
emit("    } catch {")
emit("        return $null")
emit("    }")
emit("}")
emit("")
emit("function Invoke-VaultCheckOut {")
emit("    param([string]$Repository, [string]$Path, [string]$Host)")
emit("    Write-Log \"Vault CheckOut: $Path\"")
emit("    $auth = Get-VaultAuthArgs")
emit("    if ($Host) { $auth[1] = $Host }")
emit("    & $script:VaultExe checkout @auth -repository $Repository `\"$Path`\"")
emit("    if ($LASTEXITCODE -ne 0) { throw \"Vault checkout failed: $Path\" }")
emit("}")
emit("")
emit("function Invoke-VaultGetLatest {")
emit("    param([string]$Repository, [string]$Path, [string]$LocalFolder)")
emit("    Write-Log \"Vault GetLatest: $Path\"")
emit("    $auth = Get-VaultAuthArgs")
emit("    $mappedFolder = Get-VaultWorkingFolder -Repository $Repository -ReposPath $Path")
emit("    # Omit -destpath when a working folder is already mapped and matches the target;")
emit("    # using -destpath with the same path causes \"GetToLocationOutsideWorkingFolder\" errors.")
emit("    $destArg = if ($mappedFolder -and (-not $LocalFolder -or $LocalFolder -eq $mappedFolder)) {")
emit("        @()")
emit("    } elseif ($LocalFolder) {")
emit("        @('-destpath', $LocalFolder)")
emit("    } else {")
emit("        @('-destpath', '.')")
emit("    }")
emit("    & $script:VaultExe get @auth -repository $Repository @destArg `\"$Path`\"")
emit("    if ($LASTEXITCODE -ne 0) { throw \"Vault get latest failed: $Path\" }")
emit("}")
emit("")
emit("function Invoke-VaultCheckIn {")
emit("    param([string]$Repository, [string]$Path, [string]$Comment)")
emit("    Write-Log \"Vault CheckIn: $Path\"")
emit("    $auth = Get-VaultAuthArgs")
emit("    & $script:VaultExe checkin @auth -repository $Repository -comment `\"$Comment`\" `\"$Path`\"")
emit("    if ($LASTEXITCODE -ne 0) { throw \"Vault checkin failed: $Path\" }")
emit("}")
emit("")
emit("function Invoke-VaultLabel {")
emit("    param([string]$Repository, [string]$Path, [string]$Label)")
emit("    Write-Log \"Vault Label: $Path -> $Label\"")
emit("    $auth = Get-VaultAuthArgs")
emit("    & $script:VaultExe label @auth -repository $Repository `\"$Path`\" `\"$Label`\"")
emit("}")
emit("")
emit("function Invoke-VaultUndoCheckOut {")
emit("    param([string]$Repository, [string]$Path)")
emit("    Write-Log \"Vault UndoCheckOut: $Path\"")
emit("    $auth = Get-VaultAuthArgs")
emit("    & $script:VaultExe undocheckout @auth -repository $Repository `\"$Path`\" 2>$null")
emit("}")
emit("")
emit("function Invoke-VaultCommand {")
emit("    param([string]$Repository, [string]$Command, [string]$Parameters)")
emit("    Write-Log \"Vault Custom: $Command $Parameters\"")
emit("    $auth = Get-VaultAuthArgs")
emit("    # Split parameters string respecting quoted values")
emit("    $paramArgs = @()")
emit("    if ($Parameters) {")
emit("        $paramArgs = [regex]::Matches($Parameters, '\"([^\"]*)\"|([^\\s]+)') | ForEach-Object {")
emit("            if ($_.Groups[1].Success) { $_.Groups[1].Value } else { $_.Groups[2].Value }")
emit("        }")
emit("    }")
emit("    & $script:VaultExe $Command @auth -repository $Repository @paramArgs")
emit("}")
emit("")
emit("function Invoke-VaultGetByLabel {")
emit("    param([string]$Repository, [string]$Path, [string]$Label, [string]$LocalPath)")
emit("    Write-Log \"Vault GetByLabel: $Path ($Label)\"")
emit("    $auth = Get-VaultAuthArgs")
emit("    $destArg = $LocalPath ? @('-destpath', $LocalPath) : @()")
emit("    & $script:VaultExe getlabel @auth -repository $Repository @destArg `\"$Path`\" `\"$Label`\"")
emit("}")
emit("")
emit("function Get-VaultFiles {")
emit("    param([string]$Repository, [string]$Path, [string]$Filter, [int]$TimeoutSeconds = 15)")
emit("    $auth = Get-VaultAuthArgs")
emit("    $argList = @('listfolder') + $auth + @('-repository', $Repository, $Path)")
emit("    $proc = Start-Process -FilePath $script:VaultExe -ArgumentList $argList -NoNewWindow -PassThru -RedirectStandardOutput \"$env:TEMP\\vault_list.txt\" -RedirectStandardError \"$env:TEMP\\vault_list_err.txt\"")
emit("    if (-not $proc.WaitForExit($TimeoutSeconds * 1000)) {")
emit("        $proc | Stop-Process -Force")
emit("        Write-Log \"Warning: Vault listfolder timed out after ${TimeoutSeconds}s\"")
emit("        return @()")
emit("    }")
emit("    $output = Get-Content \"$env:TEMP\\vault_list.txt\" -ErrorAction SilentlyContinue")
emit("    return $Filter ? ($output | Where-Object { $_ -like $Filter }) : $output")
emit("}")
emit("")
emit("#endregion Vault Source Control Helpers")
emit("")

# XML helpers
emit("function Get-XmlValue {")
emit("    param([string]$Path, [string]$XPath)")
emit("    if (-not (Test-Path $Path)) { return '' }")
emit("    [xml]$xml = Get-Content -Path $Path -Raw")
emit("    $node = $xml.SelectSingleNode($XPath)")
emit("    if ($node) { return $node.InnerText }")
emit("    return ''")
emit("}")
emit("")
emit("function Set-XmlValue {")
emit("    param([string]$Path, [string]$XPath, [string]$Value)")
emit("    if (-not (Test-Path $Path)) { return }")
emit("    [xml]$xml = Get-Content -Path $Path -Raw")
emit("    $node = $xml.SelectSingleNode($XPath)")
emit("    if ($node) {")
emit("        $node.InnerText = $Value")
emit("        $xml.Save($Path)")
emit("    }")
emit("}")
emit("")

# UI helpers
emit("function Confirm-Action {")
emit("    param([string]$Message, [string]$Default = 'Yes')")
emit("    Add-Type -AssemblyName System.Windows.Forms")
emit("    $result = [System.Windows.Forms.MessageBox]::Show($Message, 'Build Confirmation', 'YesNo', 'Question')")
emit("    return ($result -eq 'Yes') ? 'Yes' : 'No'")
emit("}")
emit("")
emit("function Show-RadioMenu {")
emit("    param([string]$Title, [string[]]$Options, [int]$DefaultIndex = 0)")
emit("    Add-Type -AssemblyName System.Windows.Forms")
emit("    Add-Type -AssemblyName System.Drawing")
emit("")
emit("    $font = [System.Drawing.Font]::new('Segoe UI', 9.5)")
emit("    $columns = 2")
emit("    $radioHeight = 26")
emit("    $colWidth = 280")
emit("    $padding = 12")
emit("    $rows = [Math]::Ceiling($Options.Count / $columns)")
emit("    $panelHeight = $rows * $radioHeight + $padding")
emit("    $formWidth = $colWidth * $columns + $padding * 3")
emit("")
emit("    $form = [System.Windows.Forms.Form]@{")
emit("        Text            = $Title")
emit("        StartPosition   = 'CenterScreen'")
emit("        FormBorderStyle = 'FixedDialog'")
emit("        MaximizeBox     = $false")
emit("        MinimizeBox     = $false")
emit("        TopMost         = $true")
emit("        ClientSize      = [System.Drawing.Size]::new($formWidth, $panelHeight + 50)")
emit("        Font            = $font")
emit("    }")
emit("")
emit("    $panel = [System.Windows.Forms.Panel]@{")
emit("        Location    = [System.Drawing.Point]::new($padding, $padding)")
emit("        Size        = [System.Drawing.Size]::new($formWidth - $padding * 2, $panelHeight)")
emit("        AutoScroll  = $true")
emit("    }")
emit("")
emit("    $radios = [System.Collections.Generic.List[System.Windows.Forms.RadioButton]]::new()")
emit("    for ($i = 0; $i -lt $Options.Count; $i++) {")
emit("        $col = $i % $columns")
emit("        $row = [Math]::Floor($i / $columns)")
emit("        $rb = [System.Windows.Forms.RadioButton]@{")
emit("            Text     = $Options[$i]")
emit("            Location = [System.Drawing.Point]::new($col * $colWidth + 4, $row * $radioHeight)")
emit("            Size     = [System.Drawing.Size]::new($colWidth - 8, $radioHeight)")
emit("            Checked  = ($i -eq $DefaultIndex)")
emit("            Tag      = $i")
emit("        }")
emit("        $rb.Add_DoubleClick({ $form.DialogResult = 'OK'; $form.Close() })")
emit("        $panel.Controls.Add($rb)")
emit("        $radios.Add($rb)")
emit("    }")
emit("")
emit("    $btnPanel = [System.Windows.Forms.Panel]@{")
emit("        Dock   = 'Bottom'")
emit("        Height = 40")
emit("    }")
emit("    $okButton = [System.Windows.Forms.Button]@{")
emit("        Text         = 'OK'")
emit("        DialogResult = 'OK'")
emit("        Size         = [System.Drawing.Size]::new(80, 28)")
emit("    }")
emit("    $okButton.Location = [System.Drawing.Point]::new($formWidth - 180, 6)")
emit("    $cancelButton = [System.Windows.Forms.Button]@{")
emit("        Text         = 'Cancel'")
emit("        DialogResult = 'Cancel'")
emit("        Size         = [System.Drawing.Size]::new(80, 28)")
emit("    }")
emit("    $cancelButton.Location = [System.Drawing.Point]::new($formWidth - 92, 6)")
emit("    $btnPanel.Controls.Add($okButton)")
emit("    $btnPanel.Controls.Add($cancelButton)")
emit("    $form.AcceptButton = $okButton")
emit("    $form.CancelButton = $cancelButton")
emit("")
emit("    $form.Controls.Add($panel)")
emit("    $form.Controls.Add($btnPanel)")
emit("")
emit("    if ($form.ShowDialog() -eq 'OK') {")
emit("        $selected = $radios | Where-Object { $_.Checked } | Select-Object -First 1")
emit("        return [int]$selected.Tag")
emit("    }")
emit("    return 0  # default / cancel")
emit("}")
emit("")

# MSBuild helper
emit("function Invoke-MSBuild {")
emit("    param([string]$SolutionFile, [string]$Configuration = 'Release')")
emit("    Write-Log \"Building: $SolutionFile ($Configuration)\"")
emit("    $msbuild = 'C:\\Program Files (x86)\\Microsoft Visual Studio\\2019\\Professional\\MSBuild\\Current\\Bin\\MSBuild.exe'")
emit("    if (-not (Test-Path $msbuild)) {")
emit("        $msbuild = (Get-Command msbuild -ErrorAction SilentlyContinue).Source")
emit("    }")
emit("    & $msbuild $SolutionFile /p:Configuration=$Configuration /verbosity:minimal")
emit("    if ($LASTEXITCODE -ne 0) { throw \"MSBuild failed for $SolutionFile\" }")
emit("}")
emit("")

# InstallAware helper
emit("function Invoke-InstallAware {")
emit("    param([string]$ProjectFile)")
emit("    Write-Log \"Building InstallAware: $ProjectFile\"")
emit("    & miacmd.exe $ProjectFile")
emit("    if ($LASTEXITCODE -ne 0) { throw \"InstallAware build failed for $ProjectFile\" }")
emit("}")
emit("")

# Email helper (Send-MailMessage is deprecated in PS7; using .NET SmtpClient as alternative)
emit("function Send-BuildEmail {")
emit("    param([string]$To, [string]$From, [string]$Subject, [string]$Body, [string]$SmtpServer, [int]$Port = 25)")
emit("    Write-Log \"Sending email to $To`: $Subject\"")
emit("    try {")
emit("        $message = [System.Net.Mail.MailMessage]::new($From, $To, $Subject, $Body)")
emit("        $smtp = [System.Net.Mail.SmtpClient]::new($SmtpServer, $Port)")
emit("        $smtp.Send($message)")
emit("        $smtp.Dispose()")
emit("        $message.Dispose()")
emit("    } catch {")
emit("        Write-Log \"[WARNING] Failed to send email: $_\"")
emit("    }")
emit("}")
emit("")

# Build log helpers
emit("$script:BuildLog = [System.Collections.Generic.List[string]]::new()")
emit("$script:LogFile = ''")
emit("$script:BuildTitle = ''")
emit("$script:LAST_STDOUT = ''")
emit("$script:LAST_STDERR = ''")
emit("")
emit("function Export-BuildLog {")
emit("    param([string]$Path, [string]$Mode = 'Text')")
emit("    if ($Path) { $script:BuildLog | Out-File -FilePath $Path -Force }")
emit("    return ($script:BuildLog -join \"`r`n\")")
emit("}")
emit("")

emit("#" + "=" * 70)
emit("# MAIN SCRIPT")
emit("#" + "=" * 70)
emit("")

# Process the macro
macro = root.find('Macro')
children = macro.find('Children')
if children is not None:
    for child in children:
        process_item(child)

# Write output
output_path = r'C:\Work\BuildStudio\PAApplications.ps1'
with open(output_path, 'w', encoding='utf-8') as f:
    f.write('\n'.join(out_lines))

print(f"Generated {len(out_lines)} lines of PowerShell to {output_path}")
