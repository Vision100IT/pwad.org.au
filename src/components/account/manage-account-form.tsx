import React, {FC, useCallback} from 'react';
import {useApolloClient} from '@apollo/react-hooks';
import PropTypes from 'prop-types';
import {Styled, Box, Grid, Text} from 'theme-ui';
import {Check} from 'react-feather';
import startCase from 'lodash/startCase';
import Button from '../button';
import Loading from '../loading';
import {useStripe, StripeProvider, Elements} from '../use-stripe';
import {
  MeDocument,
  CurrentSubscriptionDocument,
  StripeCheckoutSessionDocument,
  StripeCheckoutSessionMutation,
  useCancelSubscriptionMutation,
  useChangeFreeAccountMutation,
  useChangePasswordMutation,
  useCurrentSubscriptionQuery,
  User,
  StripeSubscription
} from '../queries';

const CancelSubscriptionButton: FC = () => {
  const [cancelSubscription] = useCancelSubscriptionMutation({
    update(cache, {data}) {
      cache.writeQuery({
        query: CurrentSubscriptionDocument,
        data: {
          subscription: data.cancelSubscription
        }
      });
    }
  });

  async function handleCancelSubscription(): Promise<void> {
    // eslint-disable-next-line no-alert
    if (window.confirm('Are you sure you want to cancel your subscription')) {
      await cancelSubscription();
    }
  }

  return (
    <Button
      isPrimary
      isFullWidth
      size="jumbo"
      onClick={handleCancelSubscription}
    >
      Cancel Subscription
    </Button>
  );
};

const AccountPaymentButton: FC = ({children}) => {
  const {stripe} = useStripe();
  const client = useApolloClient();

  const handleAccountPayment = useCallback(async () => {
    const {data} = await client.mutate<StripeCheckoutSessionMutation>({
      mutation: StripeCheckoutSessionDocument
    });

    const {sessionId} = data.stripeCheckoutSession;

    stripe.redirectToCheckout({
      sessionId
    });
  }, [client, stripe]);

  return (
    <Button isPrimary isFullWidth size="jumbo" onClick={handleAccountPayment}>
      {children}
    </Button>
  );
};

AccountPaymentButton.propTypes = {
  children: PropTypes.node.isRequired
};

type BuySubscriptionProps = {
  hasFreeAccount: boolean;
  changeFreeAccount: () => void;
};

const BuySubscription: FC<BuySubscriptionProps> = ({
  hasFreeAccount,
  changeFreeAccount
}) => {
  return (
    <>
      <Box sx={{width: '100%'}} marginBottom={3}>
        <Button
          isFullWidth
          iconStart={hasFreeAccount ? <Check role="img" /> : undefined}
          disabled={hasFreeAccount}
          size="jumbo"
          onClick={changeFreeAccount}
        >
          Liturgy only
        </Button>
        <Styled.p variant="prose">this option is free to use</Styled.p>
      </Box>
      <Box sx={{width: '100%'}}>
        <AccountPaymentButton>Liturgy and Music</AccountPaymentButton>
        <Styled.p variant="prose">
          this option is an Annual Subscription of $12 per year
        </Styled.p>
      </Box>
    </>
  );
};

BuySubscription.propTypes = {
  hasFreeAccount: PropTypes.bool.isRequired,
  changeFreeAccount: PropTypes.func.isRequired
};

const SubscriptionDetails: FC<StripeSubscription> = ({
  status,
  plan,
  currentPeriodEnd,
  canceledAt
}) => {
  const active = status === 'active';
  const canceled = status === 'canceled';

  return (
    <Box sx={{width: '100%'}} marginBottom={3}>
      <Text as="h3">Current Subscription</Text>
      <p>{plan}</p>

      <dl>
        <dt>Subscription status:</dt>
        <dd>{startCase(status)}</dd>

        {active && (
          <>
            <dt>Subscription end date:</dt>
            <dd>{new Date(currentPeriodEnd).toLocaleDateString()}</dd>
          </>
        )}

        {canceled && (
          <>
            <dt>Subscription canceled at:</dt>
            <dd>{new Date(canceledAt).toLocaleDateString()}</dd>
          </>
        )}
      </dl>

      {active && <CancelSubscriptionButton />}
    </Box>
  );
};

SubscriptionDetails.propTypes = {
  status: PropTypes.string.isRequired,
  plan: PropTypes.string.isRequired,
  currentPeriodEnd: PropTypes.any,
  canceledAt: PropTypes.any
};

const ManageForm: FC<User> = ({hasFreeAccount, hasPaidAccount}) => {
  const isFreeAccount = hasFreeAccount || !hasPaidAccount;

  const {data, loading} = useCurrentSubscriptionQuery();

  const [changeFreeAccount] = useChangeFreeAccountMutation({
    update(cache, {data}) {
      const {changeFreeAccount} = data;
      const {me} = cache.readQuery({query: MeDocument});
      cache.writeQuery({
        query: MeDocument,
        data: {
          me: {
            ...me,
            hasFreeAccount: changeFreeAccount.hasFreeAccount
          }
        }
      });
    }
  });

  const [changePassword] = useChangePasswordMutation();

  function handleChangeFreeAccount(): void {
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
  }

  async function handleChangePassword(): Promise<void> {
    const {data} = await changePassword();

    window.location.assign(data.changePassword.ticket);
  }

  return (
    <Grid columns={[1, 2]} gap={[3, 5]}>
      <Box sx={{width: '100%'}}>
        {loading ? (
          <Loading />
        ) : (
          <Grid columns={[1, 2]} gap={[3, 5]}>
            {data?.subscription ? (
              <SubscriptionDetails {...data.subscription} />
            ) : (
              <BuySubscription
                changeFreeAccount={handleChangeFreeAccount}
                hasFreeAccount={isFreeAccount}
              />
            )}
          </Grid>
        )}

        <Button onClick={handleChangePassword}>Change Password</Button>
      </Box>
      <Box sx={{width: '100%'}}>
        <Text as="h3">
          What is included in the Liturgy and Music account option?
        </Text>
        <Styled.p variant="prose" as="ul">
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
        </Styled.p>
      </Box>
    </Grid>
  );
};

ManageForm.propTypes = {
  hasFreeAccount: PropTypes.bool.isRequired,
  hasPaidAccount: PropTypes.bool.isRequired
};

const ManageFormProvider: FC<User> = props => (
  <StripeProvider>
    <Elements>
      <ManageForm {...props} />
    </Elements>
  </StripeProvider>
);

export default ManageFormProvider;
