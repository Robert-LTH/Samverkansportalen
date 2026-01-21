import * as React from 'react';
import { IconButton } from '@fluentui/react';
import styles from '../Samverkansportalen.module.scss';
import * as strings from 'SamverkansportalenWebPartStrings';

interface IRichTextEditorProps {
  label: string;
  value: string;
  onChange: (value: string) => void;
  disabled?: boolean;
  placeholder?: string;
}

let richTextEditorIdCounter: number = 0;
const getNextRichTextEditorId = (): string => {
  richTextEditorIdCounter += 1;
  return `richTextEditor-${richTextEditorIdCounter}`;
};

const RichTextEditor: React.FC<IRichTextEditorProps> = ({
  label,
  value,
  onChange,
  disabled,
  placeholder
}) => {
  const editorRef = React.useRef<HTMLDivElement | null>(null);
  const editorIdRef = React.useRef<string>(getNextRichTextEditorId());
  const labelId: string = `${editorIdRef.current}-label`;

  const handleInput = React.useCallback(() => {
    const nextValue: string = editorRef.current?.innerHTML ?? '';
    onChange(nextValue);
  }, [onChange]);

  const applyCommand = React.useCallback(
    (command: string): void => {
      if (disabled) {
        return;
      }

      editorRef.current?.focus();
      document.execCommand(command);
      handleInput();
    },
    [disabled, handleInput]
  );

  React.useEffect(() => {
    if (!editorRef.current) {
      return;
    }

    const currentHtml: string = editorRef.current.innerHTML;
    const nextValue: string = value ?? '';

    if (currentHtml !== nextValue) {
      editorRef.current.innerHTML = nextValue;
    }
  }, [value]);

  const toolbarButtons: { key: string; icon: string; label: string; command: string }[] = [
    { key: 'bold', icon: 'Bold', label: strings.RichTextEditorBoldButtonLabel, command: 'bold' },
    { key: 'italic', icon: 'Italic', label: strings.RichTextEditorItalicButtonLabel, command: 'italic' },
    {
      key: 'underline',
      icon: 'Underline',
      label: strings.RichTextEditorUnderlineButtonLabel,
      command: 'underline'
    },
    {
      key: 'bullets',
      icon: 'BulletedList',
      label: strings.RichTextEditorBulletListButtonLabel,
      command: 'insertUnorderedList'
    }
  ];

  return (
    <div className={styles.richTextEditor}>
      <label id={labelId} className={styles.richTextLabel} htmlFor={editorIdRef.current}>
        {label}
      </label>
      <div className={styles.richTextToolbar} role="toolbar" aria-label={label}>
        {toolbarButtons.map((button) => (
          <IconButton
            key={button.key}
            iconProps={{ iconName: button.icon }}
            title={button.label}
            ariaLabel={button.label}
            className={styles.richTextToolbarButton}
            onClick={() => applyCommand(button.command)}
            disabled={disabled}
          />
        ))}
      </div>
      <div
        id={editorIdRef.current}
        ref={editorRef}
        className={`${styles.richTextArea} ${disabled ? styles.richTextAreaDisabled : ''}`}
        role="textbox"
        aria-multiline="true"
        aria-labelledby={labelId}
        contentEditable={!disabled}
        suppressContentEditableWarning={true}
        onInput={handleInput}
        onBlur={handleInput}
        data-placeholder={placeholder}
      />
    </div>
  );
};

export default RichTextEditor;
