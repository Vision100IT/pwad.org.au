import React from 'react';
import PropTypes from 'prop-types';
import {Mutation} from 'react-apollo';
import Box from 'mineral-ui/Box';
import Flex, {FlexItem} from 'mineral-ui/Flex';
import Button from 'mineral-ui/Button';
import Text from 'mineral-ui/Text';
import TextInput from 'mineral-ui/TextInput';
import {Check} from 'react-feather';
import {string, object} from 'yup';
import {Formik, Form, FormField} from '../form';

import {CHANGE_PASSWORD, FREE_USER, ME} from '../queries';

const validationSchema = object().shape({
  password: string()
    .label('Password')
    .required(),
  newPassword: string()
    .label('New password')
    .required(),
  confirmPassword: string()
    .label('Confirm password')
    .required()
});

class ManageForm extends React.Component {
  handleChangeFreeAccount = changeFreeAccount => {
    return () =>
      changeFreeAccount({
        variables: {hasFreeAccount: true},
        optimisticResponse: {
          __typename: 'Mutation',
          changeFreeAccount: {
            __typename: 'User',
            hasFreeAccount: true
          }
        }
      });
  };

  handleChangeFreeAccountUpdate = (cache, {data}) => {
    const {changeFreeAccount} = data;
    const {me} = cache.readQuery({query: ME});
    cache.writeQuery({
      query: ME,
      data: {
        me: {
          ...me,
          hasFreeAccount: changeFreeAccount.hasFreeAccount
        }
      }
    });
  };

  handleChangePassword = changePassword => {
    return ({password, newPassword, confirmPassword}) => {
      changePassword({
        variables: {
          password,
          newPassword,
          confirmPassword
        }
      });
    };
  };

  render() {
    const {hasFreeAccount, hasPaidAccount} = this.props;
    const freeAccount = hasFreeAccount && !hasPaidAccount;

    return (
      <>
        <Flex
          gutterWidth="xxl"
          breakpoints={['narrow']}
          direction={['column-reverse', 'row-reverse']}
        >
          <FlexItem width="100%">
            <Text element="h3">
              What is included in the Liturgy and Music account option?
            </Text>
            <Text appearance="prose" element="ul">
              <li>Search Capabilities</li>
              <li>PowerPoint Slides</li>
              <li>Sound Bites</li>
              <li>Music Scores</li>
              <li>Alternate Tunes</li>
              <li>Author Index</li>
              <li>Topical Index</li>
              <li>Metre Index</li>
              <li>Occasion Index</li>
              <li>Year Index</li>
            </Text>
          </FlexItem>
          <FlexItem width="100%">
            <Flex
              breakpoints={['narrow']}
              direction={['column-reverse', 'row']}
            >
              <FlexItem width="100%">
                <Mutation
                  mutation={FREE_USER}
                  update={this.handleChangeFreeAccountUpdate}
                >
                  {changeFreeAccount => (
                    <Button
                      fullWidth
                      iconStart={freeAccount ? <Check role="img" /> : undefined}
                      disabled={hasFreeAccount || hasPaidAccount}
                      size="jumbo"
                      onClick={this.handleChangeFreeAccount(changeFreeAccount)}
                    >
                      Liturgy only
                    </Button>
                  )}
                </Mutation>
                <Text appearance="prose">this option is free to use</Text>
              </FlexItem>
              <FlexItem width="100%">
                <Button
                  primary
                  fullWidth
                  size="jumbo"
                  iconStart={hasPaidAccount ? <Check role="img" /> : undefined}
                  disabled={hasPaidAccount}
                >
                  Liturgy and Music
                </Button>
                <Text appearance="prose">
                  this option has a once off $30 fee
                </Text>
              </FlexItem>
            </Flex>
            <Mutation mutation={CHANGE_PASSWORD}>
              {changePassword => (
                <Formik
                  validationSchema={validationSchema}
                  initialValues={{
                    password: '',
                    newPassword: '',
                    confirmPassword: ''
                  }}
                  onSubmit={this.handleChangePassword(changePassword)}
                >
                  <Form>
                    <Box marginBottom="md">
                      <FormField
                        required
                        type="password"
                        label="Old password"
                        input={TextInput}
                        name="password"
                      />
                    </Box>
                    <Box marginBottom="md">
                      <FormField
                        required
                        type="password"
                        label="New password"
                        input={TextInput}
                        name="newPassword"
                      />
                    </Box>
                    <Box marginBottom="md">
                      <FormField
                        required
                        type="password"
                        label="Confirm password"
                        input={TextInput}
                        name="confirmPassword"
                      />
                    </Box>
                    <Box marginBottom="md">
                      <Button type="submit">Change Password</Button>
                    </Box>
                  </Form>
                </Formik>
              )}
            </Mutation>
          </FlexItem>
        </Flex>
      </>
    );
  }
}

ManageForm.propTypes = {
  hasFreeAccount: PropTypes.bool.isRequired,
  hasPaidAccount: PropTypes.bool.isRequired
};

export default ManageForm;