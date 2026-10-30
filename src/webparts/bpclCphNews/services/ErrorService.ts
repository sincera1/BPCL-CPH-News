export default class ErrorService {

  public static getUserMessage(error: any): string {

  if (!error) {
    return "Unexpected error occurred.";
  }

  const status =
    error?.status ||
    error?.data?.response?.status ||
    error?.response?.status;
 
  // Network
  if (status === 0) {
    return "Network error. Please check your connection.";
  }

  // Unauthorized
  if (status === 401 || status === 403) {
    return "You are not authorized to perform this action.";
  }

  // Not found
  if (status === 404) {
    return "Requested resource not found.";
  }

  // Throttling
  if (status === 429) {
    return "Too many requests. Please try again later.";
  }

  return "Something went wrong. Please contact administrator.";
}

  // Internal secure logging
  public static logError(location: string, error: any): void {

    const sanitizedError = {
      message: error?.message || "Unknown error",
      status: error?.status || "N/A",
      time: new Date().toISOString(),
      location
    }; 

    console.error("Error:", sanitizedError);

   
  }
} 