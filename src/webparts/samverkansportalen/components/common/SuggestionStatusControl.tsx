import * as React from 'react';
import { Dropdown, type IDropdownOption } from '@fluentui/react';
import styles from '../Samverkansportalen.module.scss';
import * as strings from 'SamverkansportalenWebPartStrings';
import { measureStatusDropdownWidth } from '../../utils/statusDropdown';

interface ISuggestionStatusControlProps {
  statuses: string[];
  value: string;
  isEditable: boolean;
  isDisabled: boolean;
  onChange: (status: string) => void;
}

const SuggestionStatusControl: React.FC<ISuggestionStatusControlProps> = ({
  statuses,
  value,
  isEditable,
  isDisabled,
  onChange
}) => {
  const normalizedStatuses: string[] = React.useMemo(() => {
    const seen: Set<string> = new Set();
    const items: string[] = [];

    const addStatus = (status: string | undefined): void => {
      if (!status) {
        return;
      }

      const trimmed: string = status.trim();

      if (!trimmed) {
        return;
      }

      const key: string = trimmed.toLowerCase();

      if (seen.has(key)) {
        return;
      }

      seen.add(key);
      items.push(trimmed);
    };

    statuses.forEach((status) => addStatus(status));
    addStatus(value);

    return items;
  }, [statuses, value]);

  const options: IDropdownOption[] = React.useMemo(
    () =>
      normalizedStatuses.map((status) => ({
        key: status,
        text: status
      })),
    [normalizedStatuses]
  );

  const dropdownWidth: number | undefined = React.useMemo(
    () => measureStatusDropdownWidth(normalizedStatuses),
    [normalizedStatuses]
  );
  const dropdownStyles = React.useMemo(
    () => (dropdownWidth ? { dropdown: { width: dropdownWidth } } : undefined),
    [dropdownWidth]
  );

  if (!isEditable) {
    return <span className={styles.statusBadge}>{value}</span>;
  }

  const selectedOption: IDropdownOption | undefined = options.find((option) => option.key === value);

  return (
    <Dropdown
      className={styles.statusDropdown}
      options={options}
      selectedKey={selectedOption ? selectedOption.key : value}
      onChange={(_event, option) => {
        if (!option) {
          return;
        }

        const nextStatus: string = String(option.key);

        if (nextStatus !== value) {
          onChange(nextStatus);
        }
      }}
      disabled={isDisabled}
      ariaLabel={strings.StatusLabel}
      dropdownWidth={dropdownWidth}
      styles={dropdownStyles}
    />
  );
};

export default SuggestionStatusControl;
