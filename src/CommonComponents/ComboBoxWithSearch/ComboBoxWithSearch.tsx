/* eslint-disable @typescript-eslint/explicit-function-return-type */
import * as React from "react";
import {
  ComboBox,
  IComboBox,
  IComboBoxOption,
  IComboBoxStyles,
  IStackTokens,
  Stack,
} from "@fluentui/react";

const comboBoxStyles: Partial<IComboBoxStyles> = { root: { maxWidth: 300 } };

export interface IComboBoxWithSearchProps {
  filteredOptions: IComboBoxOption[];
  onChange: (
    event: React.FormEvent<IComboBox>,
    option?: IComboBoxOption,
    index?: number,
    value?: string
  ) => void;
  label: string;
}

const stackTokens: IStackTokens = { childrenGap: 10 };

const ComboBoxWithSearch: React.FunctionComponent<IComboBoxWithSearchProps> = (
  props: IComboBoxWithSearchProps
) => {
  return (
    <Stack tokens={stackTokens}>
      <ComboBox
        label={props.label}
        options={props.filteredOptions}
        styles={comboBoxStyles}
        onChange={props.onChange}
        allowFreeform
        autoComplete="on"
        multiSelect
      />
    </Stack>
  );
};

export default ComboBoxWithSearch;
