namespace Figoint
{
    public abstract class SearchTree<T>
    {
        public abstract Leaf<T> FindNearest(float x, float y);
    }

    public class Leaf<T>
    {
        public T Data;
        public float X;
        public float Y;

        public Leaf(T data, float x, float y)
        {
            Data = data;
            X = x;
            Y = y;
        }
    }
}