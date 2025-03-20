export interface IGenerator {
  generateSlides(): Promise<void>;

  cleanup(): Promise<void>;
}
