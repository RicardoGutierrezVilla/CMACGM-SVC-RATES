const axios = require('axios');

// Set the GraphQL endpoint URL for authentication (if different, update accordingly)
const GRAPHQL_URL = process.env.GRAPHQL_URL || 'https://primefreight.betty.app/api/runtime/da93364a26fb4eeb9e56351ecec79abb';

async function login() {
  const query = `
    mutation login {
      login(
        authProfileUuid: "17838b935c5a46eebc885bae212d6d86",
        username: "agents",
        password: "admin@123"
      ) {
        jwtToken
        refreshToken
      }
    }
  `;
  
  try {
    const response = await axios.post(GRAPHQL_URL, { query });
    if (response.data.data && response.data.data.login) {
      console.log('Successfully logged in');
      return response.data.data.login;
    } else {
      throw new Error('Login failed: No login data returned');
    }
  } catch (error) {
    console.error('Error during login:', error.response?.data || error.message);
    throw error;
  }
}

module.exports = { login };


