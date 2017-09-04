package com.ctrip.framework.apollo.portal.spi.springsecurity;

import org.apache.tomcat.jdbc.pool.DataSource;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.context.annotation.Configuration;
import org.springframework.context.annotation.Profile;
import org.springframework.core.annotation.Order;
import org.springframework.security.config.annotation.authentication.builders.AuthenticationManagerBuilder;
import org.springframework.security.config.annotation.method.configuration.EnableGlobalMethodSecurity;
import org.springframework.security.config.annotation.web.builders.HttpSecurity;
import org.springframework.security.config.annotation.web.configuration.EnableWebSecurity;
import org.springframework.security.config.annotation.web.configuration.WebSecurityConfigurerAdapter;
import org.springframework.security.crypto.bcrypt.BCryptPasswordEncoder;
import org.springframework.security.crypto.password.PasswordEncoder;
import org.springframework.security.provisioning.JdbcUserDetailsManager;
import org.springframework.security.web.authentication.LoginUrlAuthenticationEntryPoint;

@Order(99)
@Profile({"auth"})
@Configuration
@EnableWebSecurity
@EnableGlobalMethodSecurity(prePostEnabled = true)
public class SpringSecurityConfigurer extends WebSecurityConfigurerAdapter {

  public static final String USER_ROLE = "user";

  @Autowired
  private DataSource datasource;

  @Override
  protected void configure(HttpSecurity http) throws Exception {
    http.csrf().disable();
    http.headers().frameOptions().sameOrigin();
    http.authorizeRequests()
        .antMatchers("/resources/**").permitAll()
        .antMatchers("/openapi/*").permitAll()
        .antMatchers("/*").hasAnyRole(USER_ROLE);
    http.formLogin().loginPage("/clogin").permitAll().failureUrl("/clogin?#/error").and().httpBasic();
    http.logout().invalidateHttpSession(true).clearAuthentication(true).logoutSuccessUrl("/clogin?#/logout");
    http.exceptionHandling().authenticationEntryPoint(new LoginUrlAuthenticationEntryPoint("/clogin"));
  }

  @Autowired
  public void configureGlobal(AuthenticationManagerBuilder auth, JdbcUserDetailsManager userDetailsService)
      throws Exception {
    PasswordEncoder encoder = new BCryptPasswordEncoder();

    auth.userDetailsService(userDetailsService).passwordEncoder(encoder);
    auth.jdbcAuthentication().dataSource(datasource).usersByUsernameQuery(
        "select username,password, enabled from users where username=?");
  }

}
