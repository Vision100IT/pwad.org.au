/** @jsx jsx */
import {FC} from 'react';
import {jsx, Styled, Box, Grid} from 'theme-ui';
import PropTypes from 'prop-types';
import Link, {childPageLinkProps} from '../link';
import {MenuItem} from '../../../queries/_types';
import {DarkTheme} from '../theme';

type FooterProps = {
  menuItems: MenuItem[];
};

const Footer: FC<FooterProps> = ({menuItems}) => {
  return (
    <DarkTheme>
      <footer
        sx={{
          variant: 'styles.footer'
        }}
      >
        <Box
          paddingTop="1rem"
          paddingBottom="5vh"
          marginX={[3, 'auto']}
          sx={{
            width: ['auto', '75%']
          }}
        >
          {menuItems && (
            <Grid columns={[1, 2, '1fr 2fr 2fr 1fr']}>
              {menuItems.map((menu) => (
                <Box key={menu._key}>
                  {menu.text && <Styled.p>{menu.text}</Styled.p>}
                  {menu.childpages && (
                    <Styled.ul>
                      {menu.childpages.map((item) => (
                        <li key={item._id}>
                          <Link {...childPageLinkProps(item)} />
                        </li>
                      ))}
                    </Styled.ul>
                  )}
                </Box>
              ))}
            </Grid>
          )}
          <Box marginY={3}>
            <Styled.p sx={{color: 'gray.4'}}>
              © 2007 - 2020 PWAD & the Presbyterian Church of Australia
            </Styled.p>
          </Box>
        </Box>
      </footer>
    </DarkTheme>
  );
};

Footer.propTypes = {
  menuItems: PropTypes.array.isRequired
};

export default Footer;
