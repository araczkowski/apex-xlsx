set define off
set verify off
set feedback off
WHENEVER SQLERROR EXIT SQL.SQLCODE ROLLBACK
begin wwv_flow.g_import_in_progress := true; end;
/

--       AAAA       PPPPP   EEEEEE  XX      XX
--      AA  AA      PP  PP  EE       XX    XX
--     AA    AA     PP  PP  EE        XX  XX
--    AAAAAAAAAA    PPPPP   EEEE       XXXX
--   AA        AA   PP      EE        XX  XX
--  AA          AA  PP      EE       XX    XX
--  AA          AA  PP      EEEEEE  XX      XX
prompt  Set Credentials...

begin

  -- Assumes you are running the script connected to SQL*Plus as the Oracle user APEX_040200 or as the owner (parsing schema) of the application.
  wwv_flow_api.set_security_group_id(p_security_group_id=>nvl(wwv_flow_application_install.get_workspace_id,1090022342582585));

end;
/

begin wwv_flow.g_import_in_progress := true; end;
/
begin

select value into wwv_flow_api.g_nls_numeric_chars from nls_session_parameters where parameter='NLS_NUMERIC_CHARACTERS';

end;

/
begin execute immediate 'alter session set nls_numeric_characters=''.,''';

end;

/
begin wwv_flow.g_browser_language := 'en'; end;
/
prompt  Check Compatibility...

begin

-- This date identifies the minimum version required to import this file.
wwv_flow_api.set_version(p_version_yyyy_mm_dd=>'2012.01.01');

end;
/

prompt  Set Application ID...

begin

   -- SET APPLICATION ID
   wwv_flow.g_flow_id := nvl(wwv_flow_application_install.get_application_id,403);
   wwv_flow_api.g_id_offset := nvl(wwv_flow_application_install.get_offset,0);
null;

end;
/

prompt  ...ui types
--

begin

null;

end;
/

prompt  ...plugins
--
--application/shared_components/plugins/process_type/pl_sviete_process_reports_to_xlsx

begin

wwv_flow_api.create_plugin (
  p_id => 90455113592055642 + wwv_flow_api.g_id_offset
 ,p_flow_id => wwv_flow.g_flow_id
 ,p_plugin_type => 'PROCESS TYPE'
 ,p_name => 'PL.SVIETE.PROCESS.REPORTS_TO_XLSX'
 ,p_display_name => 'Reports 2 XLSX'
 ,p_supported_ui_types => 'DESKTOP'
 ,p_image_prefix => '#PLUGIN_PREFIX#'
 ,p_execution_function => 'as_xlsx.create_xlsx_apex'
 ,p_substitute_attributes => true
 ,p_subscribe_plugin_settings => true
 ,p_help_text => '<div>'||unistr('\000a')||
'	Oracle APEX &quot;Process&quot; type plugin. This plugin provides the possibility to export the data from the reports on page to the xlsx (Microsoft Excel) file format.</div>'||unistr('\000a')||
'<div>'||unistr('\000a')||
'	&nbsp;</div>'||unistr('\000a')||
''
 ,p_version_identifier => '1.0'
 ,p_about_url => 'https://github.com/araczkowski/apex-xlsx/blob/master/README.md'
  );
wwv_flow_api.create_plugin_attribute (
  p_id => 90457314909221619 + wwv_flow_api.g_id_offset
 ,p_flow_id => wwv_flow.g_flow_id
 ,p_plugin_id => 90455113592055642 + wwv_flow_api.g_id_offset
 ,p_attribute_scope => 'COMPONENT'
 ,p_attribute_sequence => 1
 ,p_display_sequence => 10
 ,p_prompt => 'Show column names on first row?'
 ,p_attribute_type => 'CHECKBOX'
 ,p_is_required => false
 ,p_default_value => 'Y'
 ,p_is_translatable => true
  );
wwv_flow_api.create_plugin_attribute (
  p_id => 90458219475402656 + wwv_flow_api.g_id_offset
 ,p_flow_id => wwv_flow.g_flow_id
 ,p_plugin_id => 90455113592055642 + wwv_flow_api.g_id_offset
 ,p_attribute_scope => 'COMPONENT'
 ,p_attribute_sequence => 2
 ,p_display_sequence => 20
 ,p_prompt => 'Use region title as sheet name?'
 ,p_attribute_type => 'CHECKBOX'
 ,p_is_required => false
 ,p_default_value => 'Y'
 ,p_is_translatable => true
  );
wwv_flow_api.create_plugin_attribute (
  p_id => 90458833935425824 + wwv_flow_api.g_id_offset
 ,p_flow_id => wwv_flow.g_flow_id
 ,p_plugin_id => 90455113592055642 + wwv_flow_api.g_id_offset
 ,p_attribute_scope => 'COMPONENT'
 ,p_attribute_sequence => 3
 ,p_display_sequence => 30
 ,p_prompt => 'Region title'
 ,p_attribute_type => 'TEXT'
 ,p_is_required => false
 ,p_is_translatable => true
 ,p_help_text => 'A list with the title of the regions which should be included in the generated Excel file, separeted by a ;'||unistr('\000a')||
'If this attribute is left empty all regions will be included.'
  );
wwv_flow_api.create_plugin_attribute (
  p_id => 90460328572462075 + wwv_flow_api.g_id_offset
 ,p_flow_id => wwv_flow.g_flow_id
 ,p_plugin_id => 90455113592055642 + wwv_flow_api.g_id_offset
 ,p_attribute_scope => 'COMPONENT'
 ,p_attribute_sequence => 4
 ,p_display_sequence => 40
 ,p_prompt => 'Filename'
 ,p_attribute_type => 'SELECT LIST'
 ,p_is_required => false
 ,p_default_value => 'PAGE_NAME'
 ,p_is_translatable => true
  );
wwv_flow_api.create_plugin_attr_value (
  p_id => 90460919007468773 + wwv_flow_api.g_id_offset
 ,p_flow_id => wwv_flow.g_flow_id
 ,p_plugin_attribute_id => 90460328572462075 + wwv_flow_api.g_id_offset
 ,p_display_sequence => 10
 ,p_display_value => 'Based on the page name'
 ,p_return_value => 'PAGE_NAME'
  );
wwv_flow_api.create_plugin_attr_value (
  p_id => 90462029527481325 + wwv_flow_api.g_id_offset
 ,p_flow_id => wwv_flow.g_flow_id
 ,p_plugin_attribute_id => 90460328572462075 + wwv_flow_api.g_id_offset
 ,p_display_sequence => 20
 ,p_display_value => 'Based on the region title'
 ,p_return_value => 'REGION_NAME'
  );
wwv_flow_api.create_plugin_attr_value (
  p_id => 90462608533484704 + wwv_flow_api.g_id_offset
 ,p_flow_id => wwv_flow.g_flow_id
 ,p_plugin_attribute_id => 90460328572462075 + wwv_flow_api.g_id_offset
 ,p_display_sequence => 30
 ,p_display_value => 'User defined'
 ,p_return_value => 'USER'
  );
wwv_flow_api.create_plugin_attribute (
  p_id => 90463215937496261 + wwv_flow_api.g_id_offset
 ,p_flow_id => wwv_flow.g_flow_id
 ,p_plugin_id => 90455113592055642 + wwv_flow_api.g_id_offset
 ,p_attribute_scope => 'COMPONENT'
 ,p_attribute_sequence => 5
 ,p_display_sequence => 50
 ,p_prompt => 'User defined filename'
 ,p_attribute_type => 'TEXT'
 ,p_is_required => false
 ,p_is_translatable => true
 ,p_depending_on_attribute_id => 90460328572462075 + wwv_flow_api.g_id_offset
 ,p_depending_on_condition_type => 'EQUALS'
 ,p_depending_on_expression => 'USER'
  );
wwv_flow_api.create_plugin_attribute (
  p_id => 90463834984501783 + wwv_flow_api.g_id_offset
 ,p_flow_id => wwv_flow.g_flow_id
 ,p_plugin_id => 90455113592055642 + wwv_flow_api.g_id_offset
 ,p_attribute_scope => 'COMPONENT'
 ,p_attribute_sequence => 6
 ,p_display_sequence => 60
 ,p_prompt => 'Append sysdate to filename?'
 ,p_attribute_type => 'SELECT LIST'
 ,p_is_required => false
 ,p_default_value => 'NO'
 ,p_is_translatable => true
 ,p_help_text => 'Append sysdate to the filename of the generated Excel file. For example my_report_20101215.xml'
  );
wwv_flow_api.create_plugin_attr_value (
  p_id => 90464405679502791 + wwv_flow_api.g_id_offset
 ,p_flow_id => wwv_flow.g_flow_id
 ,p_plugin_attribute_id => 90463834984501783 + wwv_flow_api.g_id_offset
 ,p_display_sequence => 10
 ,p_display_value => 'No'
 ,p_return_value => 'NO'
  );
wwv_flow_api.create_plugin_attr_value (
  p_id => 90465024725508229 + wwv_flow_api.g_id_offset
 ,p_flow_id => wwv_flow.g_flow_id
 ,p_plugin_attribute_id => 90463834984501783 + wwv_flow_api.g_id_offset
 ,p_display_sequence => 20
 ,p_display_value => 'Yes, using format YYYYMMDD'
 ,p_return_value => 'YYYYMMDD'
  );
wwv_flow_api.create_plugin_attr_value (
  p_id => 90474231016770336 + wwv_flow_api.g_id_offset
 ,p_flow_id => wwv_flow.g_flow_id
 ,p_plugin_attribute_id => 90463834984501783 + wwv_flow_api.g_id_offset
 ,p_display_sequence => 30
 ,p_display_value => 'Yes, using format DDMMYYYY'
 ,p_return_value => 'DDMMYYYY'
  );
wwv_flow_api.create_plugin_attr_value (
  p_id => 90474805174772317 + wwv_flow_api.g_id_offset
 ,p_flow_id => wwv_flow.g_flow_id
 ,p_plugin_attribute_id => 90463834984501783 + wwv_flow_api.g_id_offset
 ,p_display_sequence => 40
 ,p_display_value => 'Yes, using format DDMM'
 ,p_return_value => 'DDMM'
  );
wwv_flow_api.create_plugin_attr_value (
  p_id => 90475411408774163 + wwv_flow_api.g_id_offset
 ,p_flow_id => wwv_flow.g_flow_id
 ,p_plugin_attribute_id => 90463834984501783 + wwv_flow_api.g_id_offset
 ,p_display_sequence => 50
 ,p_display_value => 'Yes, using format MMDD'
 ,p_return_value => 'MMDD'
  );
wwv_flow_api.create_plugin_attr_value (
  p_id => 90476021104776983 + wwv_flow_api.g_id_offset
 ,p_flow_id => wwv_flow.g_flow_id
 ,p_plugin_attribute_id => 90463834984501783 + wwv_flow_api.g_id_offset
 ,p_display_sequence => 60
 ,p_display_value => 'Yes, using format HH24MI'
 ,p_return_value => 'HH24MI'
  );
wwv_flow_api.create_plugin_attr_value (
  p_id => 90476630801779746 + wwv_flow_api.g_id_offset
 ,p_flow_id => wwv_flow.g_flow_id
 ,p_plugin_attribute_id => 90463834984501783 + wwv_flow_api.g_id_offset
 ,p_display_sequence => 70
 ,p_display_value => 'Yes, using format HH24MISS'
 ,p_return_value => 'HH24MISS'
  );
wwv_flow_api.create_plugin_attribute (
  p_id => 89081519545681297 + wwv_flow_api.g_id_offset
 ,p_flow_id => wwv_flow.g_flow_id
 ,p_plugin_id => 90455113592055642 + wwv_flow_api.g_id_offset
 ,p_attribute_scope => 'COMPONENT'
 ,p_attribute_sequence => 7
 ,p_display_sequence => 70
 ,p_prompt => 'Exclude column list'
 ,p_attribute_type => 'TEXT'
 ,p_is_required => false
 ,p_is_translatable => false
 ,p_help_text => 'exclude column list, separated by comma, for example: 1,2,8,9'
  );
wwv_flow_api.create_plugin_attribute (
  p_id => 90515930008251301 + wwv_flow_api.g_id_offset
 ,p_flow_id => wwv_flow.g_flow_id
 ,p_plugin_id => 90455113592055642 + wwv_flow_api.g_id_offset
 ,p_attribute_scope => 'COMPONENT'
 ,p_attribute_sequence => 8
 ,p_display_sequence => 80
 ,p_prompt => 'Use IR filters'
 ,p_attribute_type => 'CHECKBOX'
 ,p_is_required => false
 ,p_default_value => 'N'
 ,p_is_translatable => true
 ,p_help_text => 'Use the (named) report, search, filter and sort settings of an Interactive Report</br>'||unistr('\000a')||
'For a Classic Report only columns which have the attribute "include in Export" set are included.'
  );
null;

end;
/

commit;
begin
execute immediate 'begin sys.dbms_session.set_nls( param => ''NLS_NUMERIC_CHARACTERS'', value => '''''''' || replace(wwv_flow_api.g_nls_numeric_chars,'''''''','''''''''''') || ''''''''); end;';
end;
/
set verify on
set feedback on
set define on
prompt  ...done
 
