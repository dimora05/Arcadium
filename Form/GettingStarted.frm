<?xml version="1.0" encoding="utf-8"?>
<FORM_EXPORT name="GettingStarted">
  <DATA>
    <TABLE name="FORM">
      <RECORD>
        <FIELD name="NAME" value="GettingStarted" />
        <FIELD name="FORM_ENTITY_DEFINITION" value="" />
        <FIELD name="TYPE" value="FORM" />
        <FIELD name="DESCRIPTION" value="Simple form to test some code examples" />
        <FIELD name="TITLE" value="Getting Started" />
        <FIELD name="GROUP_ID" value="" />
        <FIELD name="MODIFIABLE" value="True" />
        <FIELD name="INCLUDE_DETAILS" value="False" />
        <FIELD name="FILE_DIRECTORY" value="smp$textfiles" />
        <FIELD name="FILE_EXTENSION" value="" />
        <FIELD name="FORM_XML" value="" />
        <FIELD name="CATEGORY" value="GENERAL" />
        <FIELD name="DEFAULT_ICON" value="INT_FORMS" />
        <FIELD name="WINDOW_STYLE" value="" />
        <FIELD name="SHOW_WELCOME_PAGE" value="False" />
        <FIELD name="HAS_ERRORS" value="False" />
      </RECORD>
    </TABLE>
  </DATA>
  <DEFINITION><![CDATA[<?xml version="1.0" encoding="utf-16"?>
<interface xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns="http://www.thermo.com/informatics/xmlns/interface/1.0">
  <forms>
    <form name="GettingStarted">
      <height>167</height>
      <width>337</width>
      <caption>GettingStarted</caption>
      <components />
      <minimumsize />
      <maximumsize />
      <controls>
        <SpinEditDefinition name="SpinEdit">
          <location>
            <x>24</x>
            <y>32</y>
          </location>
          <height>22</height>
          <width>277</width>
          <tabindex>1</tabindex>
          <propertylength>0</propertylength>
          <readonly>false</readonly>
          <BackColour />
          <ForeColour />
          <maxValue>100</maxValue>
          <minValue>1</minValue>
          <increment>1</increment>
        </SpinEditDefinition>
        <ButtonEditDefinition name="ButtonGo">
          <location>
            <x>24</x>
            <y>117</y>
          </location>
          <height>23</height>
          <width>75</width>
          <tabindex>2</tabindex>
          <caption>Go</caption>
          <imageAlign>Left</imageAlign>
          <size>
            <Width>75</Width>
            <Height>23</Height>
          </size>
        </ButtonEditDefinition>
        <ProgressBarDefinition name="ProgressBar">
          <location>
            <x>24</x>
            <y>66</y>
          </location>
          <height>23</height>
          <width>277</width>
          <style>Broken</style>
          <minimum>0</minimum>
          <maximum>100</maximum>
          <step>1</step>
          <orientation>Horizontal</orientation>
          <showpercent>false</showpercent>
          <shownumber>false</shownumber>
        </ProgressBarDefinition>
      </controls>
    </form>
  </forms>
  <Name>LayoutDefinition</Name>
</interface>]]></DEFINITION>
</FORM_EXPORT>