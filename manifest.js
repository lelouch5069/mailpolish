/*
 * MailPolish — 生成 Outlook 安装文件（manifest.xml）
 * 结构参照微软官方 Outlook 插件模板（Office-Addin-TaskPane 的 manifest.outlook.xml），
 * 只在"写邮件/回复"界面加一个按钮，打开可固定的侧边栏。
 * 网址根据当前网页所在位置自动填写，所以不需要手动改文件。
 */
(function (root) {
  'use strict';

  const ADDIN_ID = 'bd607eaa-d9c9-49dd-a555-22d964387962';
  const VERSION = '1.0.0.0';

  function xmlEscape(s) {
    return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }

  /** base: 网页所在目录，例如 https://yourname.github.io/mailpolish/ */
  function buildManifest(base) {
    let b = String(base || '').trim();
    if (!/^https:\/\//i.test(b)) throw new Error('网址必须以 https:// 开头');
    if (!b.endsWith('/')) b += '/';
    const u = path => xmlEscape(b + path);
    const origin = xmlEscape(new URL(b).origin);

    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0"
  xmlns:mailappor="http://schemas.microsoft.com/office/mailappversionoverrides/1.0" xsi:type="MailApp">
  <Id>${ADDIN_ID}</Id>
  <Version>${VERSION}</Version>
  <ProviderName>MailPolish (personal)</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="MailPolish 邮件润色"/>
  <Description DefaultValue="Checks spelling and grammar as you write and rewrites email text in a professional tone with AI."/>
  <IconUrl DefaultValue="${u('icon-64.png')}"/>
  <HighResolutionIconUrl DefaultValue="${u('icon-128.png')}"/>
  <SupportUrl DefaultValue="${u('index.html')}"/>
  <AppDomains>
    <AppDomain>${origin}</AppDomain>
  </AppDomains>
  <Hosts>
    <Host Name="Mailbox"/>
  </Hosts>
  <Requirements>
    <Sets>
      <Set Name="Mailbox" MinVersion="1.1"/>
    </Sets>
  </Requirements>
  <FormSettings>
    <Form xsi:type="ItemRead">
      <DesktopSettings>
        <SourceLocation DefaultValue="${u('index.html')}"/>
        <RequestedHeight>250</RequestedHeight>
      </DesktopSettings>
    </Form>
  </FormSettings>
  <Permissions>ReadWriteItem</Permissions>
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemIs" ItemType="Message" FormType="Read"/>
  </Rule>
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
    <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
      <Requirements>
        <bt:Sets DefaultMinVersion="1.3">
          <bt:Set Name="Mailbox"/>
        </bt:Sets>
      </Requirements>
      <Hosts>
        <Host xsi:type="MailHost">
          <DesktopFormFactor>
            <FunctionFile resid="Commands.Url"/>
            <ExtensionPoint xsi:type="MessageComposeCommandSurface">
              <OfficeTab id="TabDefault">
                <Group id="mailpolishComposeGroup">
                  <Label resid="GroupLabel"/>
                  <Control xsi:type="Button" id="mailpolishOpenPaneButton">
                    <Label resid="TaskpaneButton.Label"/>
                    <Supertip>
                      <Title resid="TaskpaneButton.Label"/>
                      <Description resid="TaskpaneButton.Tooltip"/>
                    </Supertip>
                    <Icon>
                      <bt:Image size="16" resid="Icon.16x16"/>
                      <bt:Image size="32" resid="Icon.32x32"/>
                      <bt:Image size="80" resid="Icon.80x80"/>
                    </Icon>
                    <Action xsi:type="ShowTaskpane">
                      <SourceLocation resid="Taskpane.Url"/>
                      <SupportsPinning>true</SupportsPinning>
                    </Action>
                  </Control>
                </Group>
              </OfficeTab>
            </ExtensionPoint>
          </DesktopFormFactor>
        </Host>
      </Hosts>
      <Resources>
        <bt:Images>
          <bt:Image id="Icon.16x16" DefaultValue="${u('icon-16.png')}"/>
          <bt:Image id="Icon.32x32" DefaultValue="${u('icon-32.png')}"/>
          <bt:Image id="Icon.80x80" DefaultValue="${u('icon-80.png')}"/>
        </bt:Images>
        <bt:Urls>
          <bt:Url id="Commands.Url" DefaultValue="${u('commands.html')}"/>
          <bt:Url id="Taskpane.Url" DefaultValue="${u('index.html')}"/>
        </bt:Urls>
        <bt:ShortStrings>
          <bt:String id="GroupLabel" DefaultValue="MailPolish"/>
          <bt:String id="TaskpaneButton.Label" DefaultValue="邮件润色"/>
        </bt:ShortStrings>
        <bt:LongStrings>
          <bt:String id="TaskpaneButton.Tooltip" DefaultValue="打开润色助手：边写边检查拼写和语法，一键改写成更专业的英文。"/>
        </bt:LongStrings>
      </Resources>
    </VersionOverrides>
  </VersionOverrides>
</OfficeApp>
`;
  }

  const api = { buildManifest, ADDIN_ID, VERSION };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  root.MailPolishManifest = api;
})(typeof window !== 'undefined' ? window : globalThis);
