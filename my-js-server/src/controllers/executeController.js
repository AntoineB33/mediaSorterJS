const { handleProgramLogic } = require('../services/executeService');


exports.executeProgram = async (req, res) => {
  try {
    const body = req.body;  // e.g. { functionName, timestamp, ... }
    
    // Call a service function with all the logic
    const responseData = await handleProgramLogic(body);

    // Return the response
    return res.json(responseData);
  } catch (error) {
    console.error('Error in executeProgram:', error);
    return res.status(500).json({ error: 'Internal server error' });
  }
};